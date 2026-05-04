// Local scheme-colour palette for chart series/point fills that
// reference <a:schemeClr val="accent1"> etc. instead of a direct
// <a:srgbClr val="RRGGBB">.
//
// This is a narrow helper intentionally duplicated from the drawing
// subsystem: if src/drawing/theme.ts exports `resolveSchemeColor` we
// prefer that (see chart-part.ts's dynamic import guard), otherwise
// we fall back to the defaults below. The sibling theme helper was
// not available at the time this shipped.
//
// Security: the input `val` is DOCX-controlled, so `resolveSchemeColor`
// matches it against a hard-coded allowlist of keys. Anything outside
// the allowlist returns null; the caller then falls through to the
// palette rotation in src/charts/render.ts.
//
// Colours are Office 2016 default accents plus neutrals — pre-sanitised
// so they can be emitted directly as SVG attribute values without a
// second round-trip through sanitizeCssColor.

const SCHEME_COLORS: Record<string, string> = {
    accent1: "#4472C4",
    accent2: "#ED7D31",
    accent3: "#A5A5A5",
    accent4: "#FFC000",
    accent5: "#5B9BD5",
    accent6: "#70AD47",
    dk1: "#000000",
    dk2: "#44546A",
    lt1: "#FFFFFF",
    lt2: "#E7E6E6",
    bg1: "#FFFFFF",
    bg2: "#E7E6E6",
    tx1: "#000000",
    tx2: "#44546A",
    hlink: "#0563C1",
    folHlink: "#954F72",
};

export function resolveSchemeColor(val: unknown): string | null {
    if (typeof val !== "string") return null;
    // Hard-coded allowlist lookup — no interpolation, no fallthrough.
    return Object.prototype.hasOwnProperty.call(SCHEME_COLORS, val)
        ? SCHEME_COLORS[val]
        : null;
}
