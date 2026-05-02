export function escapeClassName(className: string) {
	return className?.replace(/[ .]+/g, '-').replace(/[&]+/g, 'and').toLowerCase();
}

export function encloseFontFamily(fontFamily: string): string {
    return /^[^"'].*\s.*[^"']$/.test(fontFamily) ? `'${fontFamily}'` : fontFamily;
}

/**
 * Returns a CSS font-family value safe to interpolate into a stylesheet or
 * inline style. Strips backslashes and quotes, then wraps the identifier in
 * single quotes. Attacker-controlled DOCX strings can otherwise break out of
 * the property and inject new CSS rules (see SECURITY_REVIEW.md #4).
 * Falls back to `sans-serif` for empty / non-string input.
 */
export function sanitizeFontFamily(value: unknown): string {
    if (typeof value !== 'string') return 'sans-serif';
    const cleaned = value.replace(/["'\\;{}@<>]/g, '').trim();
    if (!cleaned) return 'sans-serif';
    return `'${cleaned}'`;
}

// Accept only pure hex-digit color values (3, 4, 6, or 8 hex digits). Anything
// else — including `red}a{}@import …` — is rejected. See SECURITY_REVIEW.md #4.
const HEX_COLOR_RE = /^[0-9A-Fa-f]{3,8}$/;
const CSS_FN_COLOR_RE = /^(rgb|rgba|hsl|hsla)\(\s*[-0-9.,%\s/deg]+\s*\)$/i;

/**
 * Returns a safe CSS color string for a DOCX-derived value, or `null` if the
 * input can't be safely emitted. Accepts bare hex (returned with a leading
 * `#`), `#hex`, and `rgb()/rgba()/hsl()/hsla()` functions with numeric args.
 * See SECURITY_REVIEW.md #4.
 */
export function sanitizeCssColor(value: unknown): string | null {
    if (typeof value !== 'string') return null;
    const v = value.trim();
    if (!v) return null;
    if (HEX_COLOR_RE.test(v)) return `#${v}`;
    if (v.startsWith('#') && HEX_COLOR_RE.test(v.slice(1))) return v;
    if (CSS_FN_COLOR_RE.test(v)) return v;
    return null;
}

// CSS-identifier fragments we interpolate into custom property names and
// selectors. Restricted to alphanum + underscore so attacker strings can't
// break out of the rule. See SECURITY_REVIEW.md #3, #4.
const SAFE_CSS_IDENT_RE = /^[A-Za-z0-9_]+$/;

export function isSafeCssIdent(value: unknown): value is string {
    return typeof value === 'string' && SAFE_CSS_IDENT_RE.test(value);
}

/**
 * Escapes a string so it can be embedded inside a CSS string literal wrapped
 * in double quotes (`content: "..."`). Backslashes are doubled, and `"` and
 * newline characters are CSS-escaped. See SECURITY_REVIEW.md #3.
 */
export function escapeCssStringContent(value: string): string {
    return value
        .replace(/\\/g, '\\\\')
        .replace(/"/g, '\\"')
        .replace(/\n/g, '\\A ')
        .replace(/\r/g, '\\D ');
}

export function splitPath(path: string): [string, string] {
    let si = path.lastIndexOf('/') + 1;
    let folder = si == 0 ? "" : path.substring(0, si);
    let fileName = si == 0 ? path : path.substring(si);

    return [folder, fileName];
}

export function resolvePath(path: string, base: string): string {
    try {
        const prefix = "http://docx/";
        const url = new URL(path, prefix + base).toString();
        return url.substring(prefix.length);
    } catch {
        return `${base}${path}`;
    }
}

// Keys that, if used as map keys on a plain `{}`, can mutate the prototype
// chain (`__proto__`) or confuse later lookups (`constructor`, `prototype`).
// DOCX strings feed these maps (comment/footnote/style ids, theme color
// names), so we filter defensively even when the downstream code happens to
// be safe. See SECURITY_REVIEW.md #6.
const UNSAFE_KEYS = new Set(['__proto__', 'constructor', 'prototype']);

export function keyBy<T = any>(array: T[], by: (x: T) => any): Record<any, T> {
    // `Object.create(null)` produces a null-prototype object: reads like
    // `m['toString']` return `undefined` rather than walking into
    // `Object.prototype`, and writes to `__proto__` set an own property
    // instead of mutating the prototype. Callers use `m[key]` / `for..in`,
    // both of which work identically on null-prototype and plain objects.
    const result: Record<any, T> = Object.create(null);
    for (const x of array) {
        const k = by(x);
        if (k == null) continue;
        const s = String(k);
        if (UNSAFE_KEYS.has(s)) continue;
        result[s] = x;
    }
    return result;
}

export function blobToBase64(blob: Blob): Promise<string> {
	return new Promise((resolve, reject) => {
		const reader = new FileReader();
		reader.onloadend = () => resolve(reader.result as string);
		reader.onerror = () => reject();
		reader.readAsDataURL(blob);
	});
}

export function isObject(item) {
    return item && typeof item === 'object' && !Array.isArray(item);
}

export function isString(item: unknown): item is string {
    return typeof item === 'string' || item instanceof String;
}

export function mergeDeep(target, ...sources) {
    if (!sources.length)
        return target;

    const source = sources.shift();

    if (isObject(target) && isObject(source)) {
        for (const key in source) {
            // Skip unsafe keys (`__proto__`, `constructor`, `prototype`) and
            // inherited properties so attacker-controlled DOCX values never
            // poison the prototype chain. See SECURITY_REVIEW.md #6.
            if (UNSAFE_KEYS.has(key)) continue;
            if (!Object.prototype.hasOwnProperty.call(source, key)) continue;

            if (isObject(source[key])) {
                const val = target[key] ?? (target[key] = {});
                mergeDeep(val, source[key]);
            } else {
                target[key] = source[key];
            }
        }
    }

    return mergeDeep(target, ...sources);
}

export function parseCssRules(text: string): Record<string, string> {
	const result: Record<string, string> = {};

	for (const rule of text.split(';')) {
		const [key, val] = rule.split(':');
		result[key] = val;
	}

	return result
}

export function formatCssRules(style: Record<string, string>): string {
	return Object.entries(style).map((k, v) => `${k}: ${v}`).join(';');
}

export function asArray<T>(val: T | T[]): T[] {
	return Array.isArray(val) ? val : [val];
}

export function clamp(val, min, max) {
    return min > val ? min : (max < val ? max : val);
}