// Scheme-colour resolution for DrawingML shapes.
//
// `<a:schemeClr val="accent1">` references the document's theme palette.
// The theme lives in `/word/theme/theme1.xml` (loaded as ThemePart) and
// maps palette slot names to 6-digit hex strings. When a document has
// no theme — or the slot isn't present — we fall back to the default
// Office 2016 palette.
//
// Security notes:
//   * Only the hard-coded allowlist of slot names (accent1..6, dk1/2,
//     lt1/2, bg1/2, tx1/2, hlink, folHlink) is accepted. Anything else
//     returns a safe default colour. DOCX-authored strings never reach
//     a CSS selector, id, or innerHTML from this module — the return
//     value is always a #RRGGBB string that sanitizeCssColor has
//     approved.
//   * Numeric modifiers (lumMod, lumOff, tint, shade, alpha) are
//     validated through xml.intAttr at the caller, then clamped here
//     before being applied in HSL space.

import { sanitizeCssColor } from '../utils';

// Default Office 2016 theme palette. Keeps rendering reasonable when
// no theme part is present.
export const DEFAULT_THEME_PALETTE: Record<string, string> = {
    accent1: '#4472C4',
    accent2: '#ED7D31',
    accent3: '#A5A5A5',
    accent4: '#FFC000',
    accent5: '#5B9BD5',
    accent6: '#70AD47',
    dk1: '#000000',
    lt1: '#FFFFFF',
    dk2: '#44546A',
    lt2: '#E7E6E6',
    bg1: '#FFFFFF',
    bg2: '#E7E6E6',
    tx1: '#000000',
    tx2: '#44546A',
    hlink: '#0563C1',
    folHlink: '#954F72',
};

// Allowlist of slot names we will look up. Unknown names return
// DEFAULT_FALLBACK_COLOUR and do not leak into an attribute.
const ALLOWED_SLOTS = new Set(Object.keys(DEFAULT_THEME_PALETTE));

// Used when a slot isn't recognised. Mid-grey is visually distinct from
// the common accent colours so a mis-rendered shape is noticeable.
const DEFAULT_FALLBACK_COLOUR = '#808080';

// ColourModifiers are collected by the parser. All values are the raw
// DOCX integers (100000 = 100%). Applied in the order declared in
// ECMA-376 §20.1.2.3.27 — alpha last.
export interface ColourModifiers {
    lumMod?: number;
    lumOff?: number;
    tint?: number;
    shade?: number;
    alpha?: number;
}

// Parsed colour reference. The renderer resolves `scheme` against the
// active palette before applying modifiers. `hex` takes precedence when
// both are present (`<a:srgbClr>` and `<a:schemeClr>` are siblings, not
// alternatives, but the srgb form wins in Word).
export interface ColourRef {
    hex?: string;             // already passed through sanitizeCssColor
    scheme?: string;          // raw slot name, validated against ALLOWED_SLOTS
    mods?: ColourModifiers;
}

// Resolves a ColourRef to a CSS colour string. Takes an optional theme
// palette (from `/word/theme/theme1.xml`); missing or malformed entries
// fall through to DEFAULT_THEME_PALETTE.
//
// Returns a 6-digit lowercase hex. Callers should pass the result to
// sanitizeCssColor (already validated here but belt-and-braces) before
// emitting to an SVG attribute.
export function resolveColour(
    ref: ColourRef | null | undefined,
    palette?: Record<string, string> | null,
): string {
    if (!ref) return DEFAULT_FALLBACK_COLOUR;

    let base: string | null = null;

    if (ref.hex) {
        // Already sanitised by the parser — trust and use directly.
        base = ref.hex;
    } else if (ref.scheme && ALLOWED_SLOTS.has(ref.scheme)) {
        const fromTheme = palette && palette[ref.scheme];
        // Theme palette entries come from DOCX — sanitise before use.
        const sanitised = sanitizeCssColor(fromTheme);
        base = sanitised ?? DEFAULT_THEME_PALETTE[ref.scheme];
    }

    if (!base) return DEFAULT_FALLBACK_COLOUR;

    const rgb = hexToRgb(base);
    if (!rgb) return DEFAULT_FALLBACK_COLOUR;

    const mods = ref.mods;
    if (!mods) return rgbToHex(rgb);

    let { r, g, b } = rgb;

    // lumMod + lumOff work in HSL space — multiply L by lumMod, add
    // lumOff. This is the most common modifier combo Word emits.
    if (typeof mods.lumMod === 'number' || typeof mods.lumOff === 'number') {
        const hsl = rgbToHsl(r, g, b);
        const mul = typeof mods.lumMod === 'number'
            ? clamp(mods.lumMod / 100000, 0, 1)
            : 1;
        const add = typeof mods.lumOff === 'number'
            ? clamp(mods.lumOff / 100000, -1, 1)
            : 0;
        hsl.l = clamp(hsl.l * mul + add, 0, 1);
        const out = hslToRgb(hsl.h, hsl.s, hsl.l);
        r = out.r; g = out.g; b = out.b;
    }

    // tint lightens toward white; shade darkens toward black. Linear
    // blend in RGB space — matches PowerPoint's observed output
    // closely enough for v1.
    if (typeof mods.tint === 'number') {
        const t = clamp(mods.tint / 100000, 0, 1);
        r = Math.round(r + (255 - r) * t);
        g = Math.round(g + (255 - g) * t);
        b = Math.round(b + (255 - b) * t);
    }
    if (typeof mods.shade === 'number') {
        const s = clamp(mods.shade / 100000, 0, 1);
        r = Math.round(r * s);
        g = Math.round(g * s);
        b = Math.round(b * s);
    }

    return rgbToHex({ r, g, b });
}

// Validates a scheme slot against the allowlist. Unknown names are
// rejected so downstream code gets a safe null and uses its fallback.
export function isAllowedSchemeSlot(val: unknown): val is string {
    return typeof val === 'string' && ALLOWED_SLOTS.has(val);
}

// Narrow helper for consumers that just want a hex from a scheme slot
// without the full ColourRef / modifier pipeline (e.g. chart series
// fills). Falls back to the default palette; unknown slots return null
// so the caller can use its own fallback (palette rotation, etc.).
export function resolveSchemeColor(val: unknown, palette?: Record<string, string>): string | null {
    if (!isAllowedSchemeSlot(val)) return null;
    const themed = palette?.[val];
    if (themed) return themed;
    return DEFAULT_THEME_PALETTE[val] ?? null;
}

// Serialises the inputs the document parser sees for a `w:color` /
// `w:shd` / `w:bdr` attribute cluster (the `themeColor` / `themeFill`
// slot name plus optional `themeTint` / `themeShade` hex bytes) into a
// compact placeholder string the renderer can feed back to
// `parseThemeColorReference`. Returns null when the slot name isn't an
// allowlisted scheme key. themeTint / themeShade are accepted only as
// 1-2 hex digits (Word's serialisation); other values are dropped.
//
// Examples:
//   buildThemeColorReference("accent1")                   → "accent1"
//   buildThemeColorReference("accent1", "66", null)       → "accent1:tint=66"
//   buildThemeColorReference("accent2", null, "80")       → "accent2:shade=80"
export function buildThemeColorReference(
    slot: string | null | undefined,
    themeTint?: string | null,
    themeShade?: string | null,
): string | null {
    if (!isAllowedSchemeSlot(slot)) return null;
    const parts: string[] = [slot];
    if (typeof themeTint === 'string' && /^[0-9a-fA-F]{1,2}$/.test(themeTint)) {
        parts.push(`tint=${themeTint.toLowerCase()}`);
    }
    if (typeof themeShade === 'string' && /^[0-9a-fA-F]{1,2}$/.test(themeShade)) {
        parts.push(`shade=${themeShade.toLowerCase()}`);
    }
    return parts.join(':');
}

// Parses a placeholder produced by `buildThemeColorReference` — or a
// raw `w:color` / `w:fill` value that may be a literal hex, `"auto"`,
// or a theme-reference form like `"accent1"` or `"accent1:tint=66"` —
// into a ColourRef the renderer can resolve at render time. Returns
// null when no usable reference is present.
//
// themeTint / themeShade are expressed as Word's single-byte hex
// (0-255). They map onto `tint` / `shade` in ColourModifiers, which
// the existing `resolveColour` pipeline applies in the normalised
// 0-100000 integer range.
//
// Security: the slot name is validated through isAllowedSchemeSlot;
// hex literals pass through sanitizeCssColor; the modifier fields
// are parsed as hex bytes and numerically clamped. Nothing from the
// input string survives as a selector / class / innerHTML sink.
export function parseThemeColorReference(raw: string | null | undefined): ColourRef | null {
    if (typeof raw !== 'string') return null;
    const trimmed = raw.trim();
    if (!trimmed || trimmed === 'auto') return null;

    // Literal hex forms ("#RRGGBB", "RRGGBB", "#RGB") — let the
    // existing sanitiser decide.
    const hexSan = sanitizeCssColor(trimmed);
    if (hexSan && /^#?[0-9a-fA-F]{3}([0-9a-fA-F]{3})?$/.test(trimmed)) {
        return { hex: hexSan };
    }

    const segments = trimmed.split(/[:\s]+/).filter(Boolean);
    if (segments.length === 0) return null;

    const slot = segments[0];
    if (!isAllowedSchemeSlot(slot)) return null;

    const ref: ColourRef = { scheme: slot };
    const mods: ColourModifiers = {};
    let gotMod = false;

    for (let i = 1; i < segments.length; i++) {
        const eq = segments[i].indexOf('=');
        if (eq < 0) continue;
        const key = segments[i].slice(0, eq);
        const val = segments[i].slice(eq + 1);
        if (!/^[0-9a-fA-F]{1,2}$/.test(val)) continue;
        const byte = parseInt(val, 16);
        if (!Number.isFinite(byte)) continue;
        // Word's themeTint / themeShade hex byte is 0-255 where 255
        // means "apply the full modifier" (i.e. 100%). Map onto the
        // 0-100000 scale that ColourModifiers expects.
        const scaled = Math.round((byte / 255) * 100000);
        if (key === 'tint') {
            // Word's themeTint byte (ECMA-376 §17.18.104): higher =
            // more tint toward white, 255 ≈ pure white, 0 = base. Map
            // directly onto ColourModifiers.tint (0-100000), which
            // resolveColour applies as r + (255-r) * tint/100000.
            mods.tint = scaled;
            gotMod = true;
        } else if (key === 'shade') {
            // Word's themeShade byte: higher = closer to base, 0 =
            // pure black. resolveColour applies `shade/100000` as a
            // multiplicative factor on each channel, so the direct
            // mapping matches.
            mods.shade = scaled;
            gotMod = true;
        }
    }
    if (gotMod) ref.mods = mods;
    return ref;
}

function clamp(v: number, lo: number, hi: number): number {
    if (v < lo) return lo;
    if (v > hi) return hi;
    return v;
}

function hexToRgb(hex: string): { r: number; g: number; b: number } | null {
    if (typeof hex !== 'string') return null;
    let h = hex.trim();
    if (h.startsWith('#')) h = h.slice(1);
    if (h.length === 3) h = h.split('').map((c) => c + c).join('');
    if (h.length !== 6) return null;
    if (!/^[0-9a-fA-F]{6}$/.test(h)) return null;
    return {
        r: parseInt(h.slice(0, 2), 16),
        g: parseInt(h.slice(2, 4), 16),
        b: parseInt(h.slice(4, 6), 16),
    };
}

function rgbToHex(rgb: { r: number; g: number; b: number }): string {
    const to2 = (n: number) => clamp(Math.round(n), 0, 255).toString(16).padStart(2, '0');
    return `#${to2(rgb.r)}${to2(rgb.g)}${to2(rgb.b)}`;
}

function rgbToHsl(r: number, g: number, b: number): { h: number; s: number; l: number } {
    const rn = r / 255, gn = g / 255, bn = b / 255;
    const max = Math.max(rn, gn, bn);
    const min = Math.min(rn, gn, bn);
    let h = 0, s = 0;
    const l = (max + min) / 2;
    if (max !== min) {
        const d = max - min;
        s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
        switch (max) {
            case rn: h = (gn - bn) / d + (gn < bn ? 6 : 0); break;
            case gn: h = (bn - rn) / d + 2; break;
            default: h = (rn - gn) / d + 4;
        }
        h /= 6;
    }
    return { h, s, l };
}

function hslToRgb(h: number, s: number, l: number): { r: number; g: number; b: number } {
    if (s === 0) {
        const v = Math.round(l * 255);
        return { r: v, g: v, b: v };
    }
    const hue2rgb = (p: number, q: number, t: number) => {
        let tt = t;
        if (tt < 0) tt += 1;
        if (tt > 1) tt -= 1;
        if (tt < 1 / 6) return p + (q - p) * 6 * tt;
        if (tt < 1 / 2) return q;
        if (tt < 2 / 3) return p + (q - p) * (2 / 3 - tt) * 6;
        return p;
    };
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    return {
        r: Math.round(hue2rgb(p, q, h + 1 / 3) * 255),
        g: Math.round(hue2rgb(p, q, h) * 255),
        b: Math.round(hue2rgb(p, q, h - 1 / 3) * 255),
    };
}
