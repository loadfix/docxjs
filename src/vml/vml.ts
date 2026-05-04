import { DocumentParser } from '../document-parser';
import { convertLength, LengthUsage } from '../document/common';
import { OpenXmlElementBase, DomType } from '../document/dom';
import xml from '../parser/xml-parser';
import { formatCssRules, parseCssRules, sanitizeCssColor } from '../utils';

// Word sometimes writes colour values with a trailing theme-colour-index
// suffix like "#4472c4 [3204]". Strip that before handing the value to
// sanitizeCssColor, which only accepts bare hex / #hex / rgb()/hsl(). See
// upstream VolodymyrBaydalka/docxjs#171 and SECURITY_REVIEW.md #4.
// Exported so the test harness can drive the sanitiser directly without
// needing a DOCX fixture that carries the `[####]` suffix.
export function sanitizeVmlColor(value: string | null | undefined): string | null {
	if (typeof value !== 'string') return null;
	const stripped = value.replace(/\s*\[\d+\]\s*$/, '');
	return sanitizeCssColor(stripped);
}

// Conservative allowlist for VML path strings before they are interpolated
// into an SVG `d=` attribute. VML path grammar uses a small set of command
// letters (m/l/c/x/e and their upper-case forms), digits, dot, minus,
// comma and whitespace. Anything outside that set is rejected — attackers
// shouldn't be able to inject `"/><script>` via a crafted path.
const SAFE_VML_PATH = /^[0-9eEmMlLcCxX.,\-\s]*$/;

// Monotonic counter for internally-generated SVG <defs> ids. DOCX-derived
// ids are never used for this purpose — see CLAUDE.md security notes.
let vmlDefsCounter = 0;
function nextVmlId(prefix: string): string {
	return `${prefix}-${++vmlDefsCounter}`;
}

export class VmlElement extends OpenXmlElementBase {
	type: DomType = DomType.VmlElement;
	tagName: string;
	cssStyleText?: string;
	attrs: Record<string, string> = {};
	wrapType?: string;
	imageHref?: {
		id: string,
		title: string
	}
}

// Small helper: build a VmlElement with a given tagName + attrs (+ children).
function makeVml(tagName: string, attrs: Record<string, string> = {}, children: VmlElement[] = []): VmlElement {
	const v = new VmlElement();
	v.tagName = tagName;
	v.attrs = attrs;
	for (const c of children) v.children.push(c);
	return v;
}

export function parseVmlElement(elem: Element, parser: DocumentParser): VmlElement {
	var result = new VmlElement();

	switch (elem.localName) {
		case "rect":
			result.tagName = "rect";
			Object.assign(result.attrs, { width: '100%', height: '100%' });
			break;

		case "oval":
			result.tagName = "ellipse";
			Object.assign(result.attrs, { cx: "50%", cy: "50%", rx: "50%", ry: "50%" });
			break;

		case "line":
			result.tagName = "line";
			break;

		case "shape":
			result.tagName = "g";
			break;

		case "textbox":
			result.tagName = "foreignObject";
			Object.assign(result.attrs, { width: '100%', height: '100%' });
			break;

		case "group":
			// Groups become a nested <svg> whose viewBox is the VML coord
			// space declared by `coordsize`/`coordorigin`. Browsers then
			// scale child coords into the CSS box derived from the group's
			// `style` (width/height on the outer <svg> element).
			result.tagName = "svg";
			applyGroupCoordSystem(elem, result);
			break;

		case "path": {
			// v:path exposes the raw VML path string via @v. Convert to an
			// SVG <path d=…> after strict character validation.
			result.tagName = "path";
			const rawPath = xml.attr(elem, "v");
			const safePath = convertVmlPathToSvg(rawPath);
			if (safePath) result.attrs.d = safePath;
			break;
		}

		// TODO: v:extrusion — 3D extrusion isn't expressible in SVG
		// without a software renderer (ray-tracer / 3D projection). We
		// deliberately skip this case rather than emit misleading 2D
		// output.
		case "extrusion":
			return null;

		default:
			return null;
	}

	for (const at of xml.attrs(elem)) {
		switch(at.localName) {
			case "style":
				// TODO(security): sanitize cssStyleText before emitting to style attribute — see SECURITY_REVIEW.md #5
				result.cssStyleText = at.value;
				break;

			case "fillcolor": {
				const fill = sanitizeVmlColor(at.value);
				if (fill) result.attrs.fill = fill;
				break;
			}

			case "from":
				const [x1, y1] = parsePoint(at.value);
				Object.assign(result.attrs, { x1, y1 });
				break;

			case "to":
				const [x2, y2] = parsePoint(at.value);
				Object.assign(result.attrs, { x2, y2 });
				break;
		}
	}

	// Child elements that live under the shape are promoted into either
	// SVG attributes (stroke/fill colour), <defs> entries (gradient/pattern
	// fills, drop shadow) or the SVG child tree.
	const defsChildren: VmlElement[] = [];

	for (const el of xml.elements(elem)) {
		switch (el.localName) {
			case "stroke":
				Object.assign(result.attrs, parseStroke(el));
				break;

			case "fill": {
				const { attrs, defs } = parseFill(el);
				Object.assign(result.attrs, attrs);
				if (defs) defsChildren.push(defs);
				break;
			}

			case "shadow": {
				const shadow = parseShadow(el);
				if (shadow) {
					defsChildren.push(shadow.defs);
					// Don't clobber a caller-provided filter — highly
					// unlikely in practice but cheap to guard against.
					if (!result.attrs.filter) {
						result.attrs.filter = `url(#${shadow.id})`;
					}
				}
				break;
			}

			case "imagedata":
				result.tagName = "image";
				Object.assign(result.attrs, { width: '100%', height: '100%' });
				result.imageHref = {
					id: xml.attr(el, "id"),
					title: xml.attr(el, "title"),
				}
				break;

			case "txbxContent":
				result.children.push(...parser.parseBodyElements(el));
				break;

			default:
				const child = parseVmlElement(el, parser);
				child && result.children.push(child);
				break;
		}
	}

	// If this is a <v:group>, rewrite each child's cssStyleText so that
	// position/size declared in VML coord units lands in the right place
	// once the browser interprets the nested <svg>'s viewBox. Without this
	// children sitting at e.g. "left:5000;top:5000" in a coordsize of
	// "10000,10000" would appear at 5000px rather than halfway across.
	if (elem.localName === "group") {
		rewriteGroupChildPositions(elem, result);
	}

	// Prepend <defs> so the url(#…) references emitted on the shape
	// itself (fill/filter attributes) resolve to them.
	if (defsChildren.length) {
		const defs = makeVml("defs", {}, defsChildren);
		result.children.unshift(defs);
	}

	return result;
}

function parseStroke(el: Element): Record<string, string> {
	const result: Record<string, string> = {
		'stroke-width': xml.lengthAttr(el, "weight", LengthUsage.Emu) ?? '1px'
	};
	const stroke = sanitizeVmlColor(xml.attr(el, "color"));
	if (stroke) result['stroke'] = stroke;
	return result;
}

// Parse <v:fill>. Returns SVG attributes to merge onto the parent shape
// plus, optionally, a <defs>-ready VmlElement (for gradient / pattern).
function parseFill(el: Element): { attrs: Record<string, string>, defs?: VmlElement } {
	const type = xml.attr(el, "type");
	const attrs: Record<string, string> = {};

	// Solid fill — the default VML behaviour. `fillcolor` on the parent
	// shape already covers this, but <v:fill color="…"/> can override.
	if (!type || type === "solid") {
		const color = sanitizeVmlColor(xml.attr(el, "color"));
		if (color) attrs.fill = color;
		return { attrs };
	}

	if (type === "gradient" || type === "gradientRadial") {
		return parseGradientFill(el, type);
	}

	if (type === "pattern" || type === "tile") {
		return parsePatternFill(el);
	}

	// Unknown fill type — be conservative and emit nothing.
	return { attrs };
}

function parseGradientFill(el: Element, type: string): { attrs: Record<string, string>, defs: VmlElement } {
	const color1 = sanitizeVmlColor(xml.attr(el, "color")) ?? "#000000";
	const color2 = sanitizeVmlColor(xml.attr(el, "color2")) ?? "#FFFFFF";
	// angle in VML is measured from the 12 o'clock position, clockwise,
	// in degrees. SVG linearGradient uses a vector from (x1,y1)→(x2,y2);
	// we compute that vector directly rather than relying on
	// gradientTransform so the output is stable across viewers.
	const rawAngle = parseFloat(xml.attr(el, "angle"));
	const angle = Number.isFinite(rawAngle) ? rawAngle : 0;
	// `focus` (-100..100) shifts the inflection point between color1 and
	// color2. For v1 we accept the value but don't implement asymmetric
	// focus handling unless it's explicitly 0 — treat missing/invalid as 0.
	const rawFocus = parseFloat(xml.attr(el, "focus"));
	const focus = Number.isFinite(rawFocus) ? clampNum(rawFocus, -100, 100) : 0;

	const id = nextVmlId("vml-grad");

	// Convert VML angle to SVG gradient endpoints on the unit square.
	// VML: 0° = bottom→top; SVG objectBoundingBox default = left→right.
	// Rotate via the gradient vector.
	const rad = (angle - 90) * Math.PI / 180; // shift so 0° points up
	const cx = 0.5, cy = 0.5;
	const x1 = cx - Math.cos(rad) * 0.5;
	const y1 = cy - Math.sin(rad) * 0.5;
	const x2 = cx + Math.cos(rad) * 0.5;
	const y2 = cy + Math.sin(rad) * 0.5;

	const stops: VmlElement[] = [];
	if (focus === 0) {
		stops.push(makeVml("stop", { offset: "0%", "stop-color": color1 }));
		stops.push(makeVml("stop", { offset: "100%", "stop-color": color2 }));
	} else {
		// Asymmetric: color1 at both ends, color2 at the focus point.
		const mid = `${(50 + focus / 2).toFixed(2)}%`;
		stops.push(makeVml("stop", { offset: "0%", "stop-color": color1 }));
		stops.push(makeVml("stop", { offset: mid, "stop-color": color2 }));
		stops.push(makeVml("stop", { offset: "100%", "stop-color": color1 }));
	}

	// gradientRadial maps to <radialGradient>; for v1 we fall through to
	// linearGradient for the common `type="gradient"` path and only emit
	// a radial gradient when explicitly requested.
	const gradientTag = type === "gradientRadial" ? "radialGradient" : "linearGradient";
	const gradAttrs: Record<string, string> = { id };
	if (gradientTag === "linearGradient") {
		Object.assign(gradAttrs, {
			x1: x1.toFixed(4),
			y1: y1.toFixed(4),
			x2: x2.toFixed(4),
			y2: y2.toFixed(4),
		});
	}

	const defs = makeVml(gradientTag, gradAttrs, stops);

	return {
		attrs: { fill: `url(#${id})` },
		defs,
	};
}

function parsePatternFill(el: Element): { attrs: Record<string, string>, defs?: VmlElement } {
	const color = sanitizeVmlColor(xml.attr(el, "color"));
	const color2 = sanitizeVmlColor(xml.attr(el, "color2"));
	// VML fills can also reference an embedded image via r:id. Loading
	// such images requires the renderer's relationship resolver, which
	// isn't accessible from the parser. For v1 we emit a swatch pattern
	// using color/color2 only; image-backed patterns degrade to the
	// solid color if we can't produce a visual substitute.
	const id = nextVmlId("vml-pat");

	if (!color && !color2) {
		// Nothing to render — fall back to no fill.
		return { attrs: {} };
	}

	// 8×8 hatched pattern using two alternating colours. Not pixel-exact
	// to Word's preset hatches but close enough for a "there is a
	// pattern here" signal at v1.
	const size = 8;
	const bg = makeVml("rect", {
		x: "0", y: "0",
		width: `${size}`, height: `${size}`,
		fill: color2 ?? "#FFFFFF",
	});
	const stripe = makeVml("path", {
		d: `M0,${size} L${size},0`,
		stroke: color ?? "#000000",
		"stroke-width": "1",
	});

	const defs = makeVml("pattern", {
		id,
		x: "0", y: "0",
		width: `${size}`, height: `${size}`,
		patternUnits: "userSpaceOnUse",
	}, [bg, stripe]);

	return {
		attrs: { fill: `url(#${id})` },
		defs,
	};
}

function parseShadow(el: Element): { id: string, defs: VmlElement } | null {
	// VML shadows are active only when on="t" / on="true". Treat absent
	// or "f" as off.
	const on = xml.attr(el, "on");
	if (on && !/^(t|true|1|on)$/i.test(on)) return null;

	const color = sanitizeVmlColor(xml.attr(el, "color")) ?? "#000000";
	const opacityRaw = xml.attr(el, "opacity");
	const opacity = parseVmlOpacity(opacityRaw);
	const [dx, dy] = parseVmlOffset(xml.attr(el, "offset"));

	const id = nextVmlId("vml-shadow");

	const feAttrs: Record<string, string> = {
		dx: dx.toFixed(2),
		dy: dy.toFixed(2),
		stdDeviation: "0",
		"flood-color": color,
		"flood-opacity": opacity.toFixed(3),
	};

	const fe = makeVml("feDropShadow", feAttrs);

	// x/y/width/height on <filter> default to a region that can clip the
	// shadow; expand it generously so offsets up to ~50% of the shape are
	// visible. `filterUnits=objectBoundingBox` is the SVG default.
	const filter = makeVml("filter", {
		id,
		x: "-50%",
		y: "-50%",
		width: "200%",
		height: "200%",
	}, [fe]);

	return { id, defs: filter };
}

// VML opacity can appear as a plain decimal ("0.5"), a fraction of 65536
// ("32768f"), or a percentage. Normalise to [0..1].
function parseVmlOpacity(val: string | null | undefined): number {
	if (!val) return 1;
	const s = val.trim();
	if (/f$/i.test(s)) {
		const n = parseFloat(s);
		return Number.isFinite(n) ? clampNum(n / 65536, 0, 1) : 1;
	}
	if (s.endsWith('%')) {
		const n = parseFloat(s);
		return Number.isFinite(n) ? clampNum(n / 100, 0, 1) : 1;
	}
	const n = parseFloat(s);
	return Number.isFinite(n) ? clampNum(n, 0, 1) : 1;
}

// VML shadow offset: `"2pt,3pt"` or `"2,3"` (unit-less = VML points).
function parseVmlOffset(val: string | null | undefined): [number, number] {
	if (!val) return [2, 2];
	const parts = val.split(',').map(p => p.trim());
	const dx = parseVmlLengthToPt(parts[0]);
	const dy = parseVmlLengthToPt(parts[1] ?? parts[0]);
	return [dx, dy];
}

function parseVmlLengthToPt(s: string | undefined): number {
	if (!s) return 0;
	const n = parseFloat(s);
	if (!Number.isFinite(n)) return 0;
	if (/pt$/i.test(s)) return n;
	if (/px$/i.test(s)) return n * 0.75;
	if (/in$/i.test(s)) return n * 72;
	if (/cm$/i.test(s)) return n * 28.3464567;
	if (/mm$/i.test(s)) return n * 2.8346457;
	// Unit-less — treat as VML points (empirically what Word emits).
	return n;
}

function parsePoint(val: string): string[] {
	return val.split(",");
}

// Convert a VML path string (used by <v:path @v="…"> and similar) into
// an SVG path `d` attribute value. The VML grammar we recognise:
//   m x,y     → moveto
//   l x,y     → lineto
//   c x1,y1 x2,y2 x,y → cubic bezier
//   x         → close path
//   e         → end path (equivalent to stop)
// Numeric values are VML EMU-equivalent; for consistency with the
// existing internal tree we scale via LengthUsage.VmlEmu.
function convertVmlPathToSvg(path: string | null | undefined): string | null {
	if (!path) return null;
	// Strict allowlist — reject anything outside the expected grammar
	// characters. This is the primary defence against injection of
	// attribute-terminators into a downstream setAttribute() call.
	if (!SAFE_VML_PATH.test(path)) return null;

	// Tokenise: letters become SVG commands directly (VML letters happen
	// to line up with SVG path commands — M/L/C/Z — via a simple map),
	// runs of numbers get converted through VmlEmu scaling.
	const cmdMap: Record<string, string> = {
		m: 'M', M: 'M',
		l: 'L', L: 'L',
		c: 'C', C: 'C',
		x: 'Z', X: 'Z',
		e: '',  E: '',
	};

	const out: string[] = [];
	const re = /([mMlLcCxXeE])|(-?\d+(?:\.\d+)?)|([,\s])/g;
	let match: RegExpExecArray | null;
	while ((match = re.exec(path)) !== null) {
		if (match[1] !== undefined) {
			const c = cmdMap[match[1]];
			if (c) out.push(c);
		} else if (match[2] !== undefined) {
			// Numeric literal: route through the same VmlEmu scaling
			// that convertPath() used historically so the path lines up
			// with the other VML coord conventions.
			out.push(convertLength(match[2], LengthUsage.VmlEmu));
		} else if (match[3] !== undefined) {
			// Preserve a separator so adjacent numbers don't fuse.
			if (out.length && !/[,\s]$/.test(out[out.length - 1])) {
				out.push(' ');
			}
		}
	}

	const joined = out.join('');
	// Collapse runs of whitespace and commas; SVG is tolerant of either
	// as separators but double-commas are invalid.
	return joined.replace(/\s+/g, ' ').replace(/\s*,\s*/g, ',').trim() || null;
}

// --- v:group coordinate handling -----------------------------------------

// Parse the group's coordsize / coordorigin attributes, compute the CSS
// dimensions from the group's own style, and set a viewBox on the nested
// <svg> that represents the group. Children drawn with raw VML
// coordinates then land in the right CSS location.
function applyGroupCoordSystem(elem: Element, result: VmlElement): void {
	const [csx, csy] = parseCoordPair(xml.attr(elem, "coordsize")) ?? [1000, 1000];
	const [cox, coy] = parseCoordPair(xml.attr(elem, "coordorigin")) ?? [0, 0];

	// viewBox is "min-x min-y width height".
	result.attrs.viewBox = `${cox} ${coy} ${csx} ${csy}`;
	result.attrs.preserveAspectRatio = "none";

	// Stash the scale for child rewriting. We don't send raw numbers as
	// attrs so they don't end up on the SVG element.
	(result as any).__groupCoord = { csx, csy, cox, coy };
}

// VML group children declare their own left/top/width/height in the
// group's VML coord space via CSS — e.g. style="position:absolute;
// left:100;top:200;width:500;height:500". The renderer drops a child's
// cssStyleText, so we fold those values into explicit x/y/width/height
// SVG attributes here, where the viewBox from applyGroupCoordSystem
// gives them the right visual placement.
function rewriteGroupChildPositions(_groupElem: Element, group: VmlElement): void {
	for (const child of group.children) {
		if (!(child instanceof VmlElement)) continue;
		const style = child.cssStyleText;
		if (!style) continue;

		const rules = parseCssRules(style);
		// parseCssRules tolerates a trailing empty segment; clean up undef/empty keys.
		const left = parsePositionValue(rules.left);
		const top = parsePositionValue(rules.top);
		const width = parsePositionValue(rules.width);
		const height = parsePositionValue(rules.height);

		// Leaf shapes that want x/y/width/height (rect, image, ellipse as
		// cx/cy/r) — set them only when the tag expects those attrs.
		switch (child.tagName) {
			case "rect":
			case "image":
			case "foreignObject":
			case "svg":
				if (left != null) child.attrs.x = left.toString();
				if (top != null) child.attrs.y = top.toString();
				if (width != null) child.attrs.width = width.toString();
				if (height != null) child.attrs.height = height.toString();
				break;
			case "ellipse":
				if (left != null && width != null) {
					child.attrs.cx = (left + width / 2).toString();
					child.attrs.rx = (width / 2).toString();
				}
				if (top != null && height != null) {
					child.attrs.cy = (top + height / 2).toString();
					child.attrs.ry = (height / 2).toString();
				}
				break;
			case "g":
				// Wrap <g> children in a transform so their internal
				// coords (typically "0 0 width height") land at the
				// desired offset.
				if (left != null || top != null) {
					const tx = left ?? 0;
					const ty = top ?? 0;
					const existing = child.attrs.transform ?? '';
					child.attrs.transform = `translate(${tx} ${ty}) ${existing}`.trim();
				}
				break;
			// line/path/etc. already carry absolute coords in their own
			// attrs, so cssStyleText-derived position doesn't apply.
		}
	}
}

// Parse a "1000,2000" pair with a conservative numeric cast.
function parseCoordPair(val: string | null | undefined): [number, number] | null {
	if (!val) return null;
	const parts = val.split(',').map(s => parseFloat(s.trim()));
	if (parts.length !== 2 || !Number.isFinite(parts[0]) || !Number.isFinite(parts[1])) return null;
	return [parts[0], parts[1]];
}

// Parse a CSS length value from inside a VML group child. These are
// normally unit-less (interpreted as VML coord units) but may carry
// "pt"/"px" — strip the unit and fall back to a plain number. Invalid
// input → null (caller skips the attr entirely rather than emitting
// NaN into an SVG attribute).
function parsePositionValue(val: string | undefined): number | null {
	if (val == null) return null;
	const n = parseFloat(val);
	return Number.isFinite(n) ? n : null;
}

function clampNum(val: number, min: number, max: number): number {
	return val < min ? min : val > max ? max : val;
}
