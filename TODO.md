# Upstream issue triage

Deep review of all 51 open issues + 4 open PRs on the upstream source repo
[`VolodymyrBaydalka/docxjs`](https://github.com/VolodymyrBaydalka/docxjs/issues)
originally triaged **2026-05-02**; re-audited against master on **2026-05-07**
(HEAD `6479363`). Each item was compared against the current state of this fork
(master) and its commit history — waves 1–8 shipped ~90 fidelity items that
the original triage flagged as missing. This file now lists only what remains.

See the "Genuinely remaining" section at the bottom for a bounded, verified
list of open work.

---

## Highest-impact items

### [#39 — Page break](https://github.com/VolodymyrBaydalka/docxjs/issues/39) (resolved-in-fork, high)

The longest-standing upstream complaint. Fixed here: `experimentalPageBreaks`
option in [`src/page-break.ts`](src/page-break.ts) + per-page thumbnails in
[`src/thumbnails.ts`](src/thumbnails.ts). Both follow-up blockers — repeated
headers/footers on every sub-page and mid-element splitting (table row
boundaries) — are now shipped (PR #44). The mid-paragraph split case remains
out of scope. Worth promoting off "experimental" once consumers have had a
release cycle to validate.

### [#194 — PR: sanitize hyperlink URIs to prevent XSS](https://github.com/VolodymyrBaydalka/docxjs/pull/194) (resolved-in-fork; high)

PR #23 landed the fix with a stricter allowlist (`http`/`https`/`mailto`/`tel`/`ftp(s)`
plus fragments/relatives) and an exported `isSafeHyperlinkHref` helper.
See [`SECURITY_REVIEW.md`](SECURITY_REVIEW.md) finding #2.

---

## New features — status as of 2026-05-07

| # | Title | Status |
|---|---|---|
| [#85](https://github.com/VolodymyrBaydalka/docxjs/issues/85) | Comments support | **Shipped** — full sidebar, anchored + packed layouts, CSS Highlight API, PR #2 + follow-ups |
| [#167](https://github.com/VolodymyrBaydalka/docxjs/issues/167) | DrawingML shapes | **Shipped** — `parseDrawingShape` / `renderShape` (Wave 3, PR #50) |
| [#124](https://github.com/VolodymyrBaydalka/docxjs/issues/124) | Image alt tags | **Shipped** — `wp:docPr@descr` + `pic:cNvPr@descr` + `a:blip@descr` propagated to `<img alt>` (Wave 1, PR #48) |
| [#128](https://github.com/VolodymyrBaydalka/docxjs/issues/128) | `w14:paraId` on paragraphs | **Shipped** — emitted as `data-para-id` (PR #30) |
| [#57](https://github.com/VolodymyrBaydalka/docxjs/issues/57) / [#91](https://github.com/VolodymyrBaydalka/docxjs/issues/91) | Chart support | **Shipped** — classic chart renderer (PR #51) + ChartEx sunburst / treemap / waterfall / funnel / histogram (Waves 4–8) |
| [#61](https://github.com/VolodymyrBaydalka/docxjs/issues/61) | Shape support | **Shipped** — same pipeline as #167 |
| [#88](https://github.com/VolodymyrBaydalka/docxjs/issues/88) | Link / SVG / chart | **Shipped** — covered by the charts + shapes work; SVG images already rendered |
| [#4](https://github.com/VolodymyrBaydalka/docxjs/issues/4) | Table of contents | **Partial** — field results render (cached `\r` text is emitted; REF / PAGEREF / HYPERLINK produce clickable `<a href="#...">`). A true live TOC (parsing `TOC \o "1-3"` and regenerating from heading outline) is out of scope for a read-only viewer. |
| [#12](https://github.com/VolodymyrBaydalka/docxjs/issues/12) | CSP-strict mode | **Open** — no nonce threading yet; inline `<style>` + `style` attrs still injected. |

---

## Bugs to investigate

Ordered by priority hint. These are the upstream issues that are still open
against the fork.

### Medium priority

| # | Title | Where to look |
|---|---|---|
| [#54](https://github.com/VolodymyrBaydalka/docxjs/issues/54) | Tables extend into right padding of main section | `html-renderer.ts` `renderTable` + table-width math |
| [#76](https://github.com/VolodymyrBaydalka/docxjs/issues/76) | Docx-embedded header/footer render issue | `src/header-footer/` parts |
| [#174](https://github.com/VolodymyrBaydalka/docxjs/issues/174) | Table background / text colour not rendered correctly | Theme color resolution |
| [#181](https://github.com/VolodymyrBaydalka/docxjs/issues/181) | Numbered lists continue across multiple previews | Shared global counters when mounting multiple renderers |
| [#183](https://github.com/VolodymyrBaydalka/docxjs/issues/183) | Source map issue causes CI/CD failures | `dist/*.map` has `sourcesContent: null` + missing src paths |
| [#195](https://github.com/VolodymyrBaydalka/docxjs/issues/195) | text-indent incorrect | `firstLineChars` ignored |

Fixed since last triage:
- #80 / #97 (text box display) — shipped via DrawingML `wps:txbx` (Wave 3).
- #102 / #130 / #155–159 — shipped via `wrapSquare` / `wrapTight` / `wrapThrough` (Wave 1).
- #51 — WMF/EMF decoder shipped (Wave 7).
- #59 — wrap-mode regressions handled as part of the same Wave 1 work.

### Low priority

| # | Title | Notes |
|---|---|---|
| [#161](https://github.com/VolodymyrBaydalka/docxjs/issues/161) | `&` appears with curly or square brackets | XML entity escaping in a specific path |
| [#178](https://github.com/VolodymyrBaydalka/docxjs/issues/178) | tabStops display bug | Decimal / bar tab styles still unhandled (see `src/javascript.ts:51`) |
| [#187](https://github.com/VolodymyrBaydalka/docxjs/issues/187) | Empty `<p>` has wrong height | Empty paragraph placeholder |

---

## Resolved in fork

Close out on upstream with a link to our commits / PRs. PR numbers are in
this fork's local repo.

| # | Title | Fix in fork |
|---|---|---|
| [#39](https://github.com/VolodymyrBaydalka/docxjs/issues/39) | Page break | `experimentalPageBreaks` (PR #24); headers/footers + table row splits (PR #44) |
| [#51](https://github.com/VolodymyrBaydalka/docxjs/issues/51) | WMF rendering | Wave 7 WMF/EMF decoder (`src/common/vector-image.ts`, PR #55) |
| [#57](https://github.com/VolodymyrBaydalka/docxjs/issues/57) / [#91](https://github.com/VolodymyrBaydalka/docxjs/issues/91) | Charts | Wave 4 chart renderer (PR #51); Waves 6–8 ChartEx (PRs #54, #55, #57) |
| [#59](https://github.com/VolodymyrBaydalka/docxjs/issues/59) / [#102](https://github.com/VolodymyrBaydalka/docxjs/issues/102) / [#130](https://github.com/VolodymyrBaydalka/docxjs/issues/130) / [#155–159](https://github.com/VolodymyrBaydalka/docxjs/issues/155) | Floating image wrap | Wave 1 wrap modes (`parseDrawingWrapper`, PR #48) |
| [#61](https://github.com/VolodymyrBaydalka/docxjs/issues/61) / [#167](https://github.com/VolodymyrBaydalka/docxjs/issues/167) / [#80](https://github.com/VolodymyrBaydalka/docxjs/issues/80) / [#97](https://github.com/VolodymyrBaydalka/docxjs/issues/97) | DrawingML shapes + text boxes | Wave 3 shapes pipeline (PR #50) |
| [#65](https://github.com/VolodymyrBaydalka/docxjs/pull/65) (PR) | SDT checkbox | Wave 4 SDT controls (PR #51) — independent implementation with wider SDT coverage |
| [#85](https://github.com/VolodymyrBaydalka/docxjs/issues/85) | Comments support | Full sidebar feature (PR #2 + follow-ups) |
| [#117](https://github.com/VolodymyrBaydalka/docxjs/issues/117) | `WordDocument` typings | `parseAsync` / `renderAsync` return `Promise<WordDocument>` (PR #27) |
| [#124](https://github.com/VolodymyrBaydalka/docxjs/issues/124) | Image alt tags | Wave 1 (PR #48) |
| [#128](https://github.com/VolodymyrBaydalka/docxjs/issues/128) | `w14:paraId` emission | PR #30 |
| [#138](https://github.com/VolodymyrBaydalka/docxjs/issues/138) | Negative `<svg>` width | Clamp + viewBox (PR #43) |
| [#171](https://github.com/VolodymyrBaydalka/docxjs/issues/171) | `#4472c4 [3204]` colour | `sanitizeVmlColor` (PR #29) |
| [#176](https://github.com/VolodymyrBaydalka/docxjs/issues/176) | Auto-pagination | Same as #39 |
| [#179](https://github.com/VolodymyrBaydalka/docxjs/issues/179) | Clickable hyperlinks | `renderHyperlink` + allowlist (PR #23) |
| [#188](https://github.com/VolodymyrBaydalka/docxjs/issues/188) | `renderAsync` ESM export | Fixed in build config |
| [#189](https://github.com/VolodymyrBaydalka/docxjs/pull/189) (PR) | Chart rendering | Independent chart renderer shipped (Wave 4, PR #51) — broader coverage than the upstream PR; not adopted verbatim |
| [#194](https://github.com/VolodymyrBaydalka/docxjs/pull/194) (PR) | URI allowlist | PR #23 |
| [#196](https://github.com/VolodymyrBaydalka/docxjs/issues/196) | `classNameOfCnfStyle` throw | Null-guard (PR #31) |

---

## Not applicable

Upstream-specific toolchain issues that don't apply to our fork.

| # | Title | Why |
|---|---|---|
| [#122](https://github.com/VolodymyrBaydalka/docxjs/issues/122) | Webpack `Module not found` | Legacy webpack4 toolchain |
| [#127](https://github.com/VolodymyrBaydalka/docxjs/issues/127) | React 17 | Packaging/toolchain |
| [#144](https://github.com/VolodymyrBaydalka/docxjs/issues/144) | React 16 bundler | Same |

---

## Needs more info

| # | Title |
|---|---|
| [#116](https://github.com/VolodymyrBaydalka/docxjs/issues/116) | Header image + TOC styling |
| [#123](https://github.com/VolodymyrBaydalka/docxjs/issues/123) | Bulleted list |
| [#133](https://github.com/VolodymyrBaydalka/docxjs/issues/133) | Header/footer not displaying |
| [#175](https://github.com/VolodymyrBaydalka/docxjs/issues/175) | `OpenXmlPackage#load` in micro-frontend |

---

## Duplicate / stale

| # | Title | Why |
|---|---|---|
| [#38](https://github.com/VolodymyrBaydalka/docxjs/issues/38) | jszip "end of central directory" | Corrupted input, not a docxjs bug |
| [#88](https://github.com/VolodymyrBaydalka/docxjs/issues/88) | Link/SVG/chart request | Covered by #57/#61/#167 |
| [#91](https://github.com/VolodymyrBaydalka/docxjs/issues/91) | Chart content not displayed | Duplicate of #57 |
| [#111](https://github.com/VolodymyrBaydalka/docxjs/issues/111) | TIFF images | Upstream wontfix; demo has a TIFF preprocessor |
| [#125](https://github.com/VolodymyrBaydalka/docxjs/issues/125) | `This dependency was not found` | Toolchain-specific |
| [#129](https://github.com/VolodymyrBaydalka/docxjs/issues/129) | ENOENT on source maps | Related to #183 |
| [#162](https://github.com/VolodymyrBaydalka/docxjs/issues/162) | Load document programmatically | Usage question |
| [#168](https://github.com/VolodymyrBaydalka/docxjs/issues/168) | Contact info | No actionable content |
| [#192](https://github.com/VolodymyrBaydalka/docxjs/pull/192) (PR) | README typo | Upstream-only content |

---

## Missing Word features (read-only) — residual

Audit conducted 2026-05-02 by comparing the renderer against ECMA-376 /
WordprocessingML read-only fidelity, and re-walked 2026-05-07 against master
at HEAD `6479363`. The original audit listed ~90 entries across seven
subsections; Waves 1–8 shipped most of them. The tables below list **only
items still unimplemented or partial**. For every item below, the source
pointer has been verified against the current tree.

### Text and typography

Run-level formatting now covered: `b`, `i`, `u`, `strike`, `dstrike`, `caps`,
`smallCaps`, `color`, `sz`, `position`, `rFonts`, `highlight`, `shd`,
`vertAlign`, `bdr`, `kern`, `w` (scale), `spacing` (tracking), `emboss`,
`imprint`, `outline`, `shadow`, `em`, `ruby`, `fitText`, `cs` / `bCs` /
`iCs` / `szCs`, `bdo`, `rtl`, `softHyphen`, `cr`, `vanish` / `specVanish`.

| Item | Status | Source pointer |
|---|---|---|
| `w:lang` on runs / paragraphs → HTML `lang` attribute | Not implemented | Parsed as `$lang` (`document-parser.ts:3001`) but renderer only applies `lang` on `<ruby>` (`html-renderer.ts:2482`). Accessibility + browser hyphenation still miss per-run language. |
| Font substitution chain (`w:altName`, `w:panose1`, `w:sig`) | Partial | `altName` parsed in `font-table/fonts.ts:39`; `panose1` / `sig` ignored. Embedded fonts from `fontTable.xml` load but no fallback chain is constructed. |
| `w:vertAlign val="baseline"` reset inside a superscript ancestor | Gap | Nested runs that explicitly reset back to baseline aren't unwrapped from `<sub>/<sup>`. |

### Paragraphs and sections

Now covered: `widowControl`, `keepNext`, `keepLines`, `pageBreakBefore` (all
emit CSS `break-before/inside/after` — `html-renderer.ts:2307`), `pgBorders`
(`html-renderer.ts:685`), `lnNumType`, `docGrid`, `mirrorMargins`, per-col
widths (`html-renderer.ts:770`), outline level → `<h1>..<h6>`
(`getHeadingTagName`, `html-renderer.ts:254`), `w:num` `lvlOverride` /
`startOverride` (`document-parser.ts:497`), frame drop caps with `lines`
support.

| Item | Status | Source pointer |
|---|---|---|
| `w:pBdr` art borders (`apples`, `stars`, etc.) | Gap | `parseBorderProperties` only knows left/right/top/bottom; DrawingML art border values fall through. |
| Tab stops — decimal, bar alignment + custom leaders beyond dot/hyphen | Partial | `javascript.ts:58` handles right/centre/dot/middleDot/hyphen/heavy/underscore; `decimal` and `bar` fall through. Tab engine only runs when `options.experimental` is truthy. |
| Numbering: `w:lvlRestart` mid-document | Partial | Parser reads `lvlRestart` / `isLgl` / `lvlJc` (`document-parser.ts:626`); rendering restart logic uses CSS `counter-set` on level change, which doesn't honour an explicit `w:lvlRestart` between same-level paragraphs. |
| Watermarks (`behindDoc` anchor) | Gap | `behindDoc` is parsed (`document-parser.ts:1407`) but never applied to `z-index`; the shape renders in-flow rather than layered under the body. |
| Footnote/endnote physical-bottom placement + `numRestart` policy | Partial | Footnotes append at section end (`html-renderer.ts:334` has a TODO). Numbering-restart policies (`numRestart` in `footnotePr`) not parsed. |
| `w:settings evenAndOddHeaders` | Partial | `evenAndOddHeaders` now parsed in `settings/settings.ts:29`, but `renderHeaderFooter` doesn't gate the even-header pick on the setting — even header fires for placeholder definitions. |

### Tables

Now covered: diagonal borders (`tl2br` / `tr2bl` — `html-renderer.ts:3643`),
`tblHeader` → `<thead>` with `<th scope="col">` (`html-renderer.ts:3320`),
`cantSplit`, `noWrap`, `tblpPr` with `horzAnchor` / `vertAnchor` / `tblpX` /
`tblpY`, numeric `tblStyleColBandSize` / `tblStyleRowBandSize` via
`data-band`, `w:shd` pattern values (pct / diagStripe / horzStripe etc.,
`document-parser.ts:3356`).

| Item | Status | Source pointer |
|---|---|---|
| `w:tblHeader` repeating rows across a *paginated* break | Partial | Semantic `<thead>` + `<th scope="col">` emitted; visual pagination (`page-break.ts`) duplicates header rows when splitting an oversized table, but multi-page header repetition at the browser level isn't engaged. |
| `w:tblBorders` inside a `tblStylePr` (first-row borders only) | Gap | Conditional sub-style `tblBorders` attach to the whole selector and can over-border sibling cells. |
| Nested tables | Gap | Supported structurally; CSS inheritance from outer `tblStyle` cascades into nested tables beyond what Word does. |

### Graphics (DrawingML / VML)

Now covered: DrawingML shapes `wps:wsp` + groups `wpg:wgp` + text boxes
`wps:txbx` + custom geometry `a:custGeom`, glow / reflection / shadow /
softEdge effects (`drawing/shapes.ts`), `wrapSquare` / `wrapTight` /
`wrapThrough` + polygon shape-outside (`document-parser.ts:1471`),
`relativeFrom="paragraph"` / `"column"` absolute/float handling, `distT/B/L/R`
padding, `xfrm@rot` image rotation, `a:srcRect` crop, WMF/EMF decoder
(`common/vector-image.ts`, Wave 7), `parseThemeColorReference` +
`resolveSchemeColor` + `lumMod` / `lumOff` theme resolution (`drawing/theme.ts`),
VML `path` / `shadow`, VML gradient + pattern fills, VML group
`coordsize` / `coordorigin`, classic charts + ChartEx (sunburst / treemap /
waterfall / funnel / histogram), SmartArt placeholder + mc:Fallback
preferred path.

| Item | Status | Source pointer |
|---|---|---|
| SmartArt layout engine (list / cycle / hierarchy / pyramid / matrix rendered from `dgm:` data) | Partial | `parseSmartArtReference` (`document-parser.ts:1655`) unwraps `<mc:Fallback>` where present — covers most real docs. No layout engine for fallback-less SmartArt. |
| Image effects: `a:alphaModFix` (transparency), `a:duotone`, `a:lum`, `a:biLevel` on pictures | Not implemented | No DrawingML effect pipeline for pictures. (Shape effects are covered via `drawing/shapes.ts`.) |
| Image `a:tile` variants beyond default | Partial | `a:tile` is handled in `vml/vml.ts:233` for VML fills; DrawingML-side `a:tile` on pictures isn't honoured. |
| OLE objects (`w:object` / `o:OLEObject`) | Not implemented | No `object` branch in `parseRun`. Legacy Equation Editor 3.x embeds, embedded spreadsheets etc. are dropped. |
| Equations: `w:oMath` `limLoc` respected in render | Partial | Parsed (`document-parser.ts:1280`) but nary with under/over limits always renders as `<msubsup>`. |

### Fields and content

Now covered: `w:fldSimple` + complex fields (`renderSimpleField` /
`parseFieldInstruction`, `html-renderer.ts:1742`), cross-references (REF /
PAGEREF rendered as `<a href="#anchor">`, `html-renderer.ts:2169`),
HYPERLINK fields with tooltip + target frame, SDT checkbox / dropdown /
combo / date / gallery / docPartList (`document-parser.ts:691`,
`html-renderer.ts:2529`), `w:tooltip` on hyperlinks (`document-parser.ts:977`,
`html-renderer.ts:2433`), bookmark `colFirst`/`colLast` column ranges
(`html-renderer.ts:3458`), `<mc:AlternateContent>` fallback
(`document-parser.ts:1343`).

| Item | Status | Source pointer |
|---|---|---|
| Form fields (`FORMTEXT`, `FORMCHECKBOX`, `FORMDROPDOWN`) — legacy `w:fldChar ffData` | Not implemented | The newer SDT-based controls render; the legacy form-field path is not handled. |
| Glossary documents (`/word/glossary/document.xml`) | Not implemented | Not loaded by `word-document.ts`. |
| Data-bound `w:sdt` pulling from custom XML parts | Not implemented | Custom XML parts aren't loaded. |
| `w:altChunk` embedded HTML / RTF content | Not implemented (intentional) | Renderer returns `null` (`html-renderer.ts:1919`) — removed for security (see SECURITY_REVIEW.md #1). |

### Document-level

| Item | Status | Source pointer |
|---|---|---|
| Document protection markers (`w:documentProtection`) | Not implemented | Not parsed; no visual affordance. |
| Compatibility settings (`w:compat`) | Not implemented | Not parsed; rendering always uses modern line-height defaults. |
| W3C XML digital signatures | Not implemented | Not applicable to read-only rendering beyond a "signed" badge. |

### Accessibility and semantics

Now covered: heading outline (`<h1>..<h6>`), `role="document"` wrapper, ARIA
landmarks, first-row cells get `scope="col"`, table header rows → `<thead>`
with `<th scope="col">`, image alt text from `wp:docPr@descr` + fallbacks,
SDT `alias`/`tag` → `aria-label`, shared `oox-*` class family alongside
`docx-*`.

| Item | Status | Source pointer |
|---|---|---|
| Run / paragraph `w:lang` → HTML `lang` attribute | Not implemented | Same entry as in "Text and typography" — parsed as `$lang` but not emitted on runs / paragraphs. Blocks browser hyphenation + screen-reader language detection. |
| `w:style.uiPriority` / `w:hidden` used to hide auto-TOC rows | Not implemented | Parsed but ignored at `document-parser.ts:341`. |

---

## Conformance gaps (2026-05-04 overnight corpus run)

All 16 gaps surfaced by the `loadfix/ooxml-validate` 950-case run against
`26c54dd` are closed in branch `fix/w1-f-conformance-gaps-2026-05-04` and
merged as `d6ab27d`. Key library-side changes:

- **`valueOfBorder()`** — interpolated a missing `@w:color` as `"null"`,
  producing an unparseable shorthand. Fallback to `autos.borderColor` when
  colour is null or "auto".
- **`renderBreak`** — now emits `<br class="docx-page-break page-break"
  data-page-break>` regardless of `experimentalPageBreaks`.
- **`createPageElement`** — writes `data-page-orientation` /
  `data-orientation` / `class="page-landscape|page-portrait"` on the
  `<section>` whenever `props.pageSize` is known.
- **`parseTableRowProperties`** — presence-only `<w:tblHeader/>` now lands
  as `isHeader=true` (default-true in `boolAttr`), so `<thead>` +
  `<th scope="col">` emit.
- **`renderComments*`** — always emit a minimal hook (`<sup
  class="docx-comment-ref comment-reference" data-comment="...">`) even
  when `options.renderComments === false`.
- **`renderFootnoteReference`** — emits `class="docx-footnote-ref
  footnote-ref footnote"` + both `data-footnote-id` and `data-footnote`.
- **`wrapFieldResult`** — writes `data-field="REF"` + `.field-ref` /
  `.docx-field-ref` on the `<a>` wrapping a REF / PAGEREF cached result.
- **Paragraph alignment** — `5907b76` added the `distribute →
  text-align: justify` mapping; paragraph styles are emitted inline via
  `cssStyle`, so conformance selectors keyed on computed style match.

Subsequent W9 waves added a11y landmarks, `w:lang` fallthrough / `tblLook`
header-row fixes, responsive render option, and the shared `oox-*` CSS
class family (commits `071b3dd`, `832d345`, `fc0595f`, `4e39d1b`).

---

## Genuinely remaining (as of 2026-05-07, HEAD `6479363`)

After Waves 1–8 the remaining work is small. Items verified as NOT yet
shipped by grepping master:

### Bugs (from upstream issue triage)

- **#54** — tables extend into right padding of main section.
- **#76** — docx-embedded header/footer render issue (needs a repro).
- **#161** — `&` appears with curly / square brackets (XML entity escaping
  on a specific path).
- **#174** — table background / text colour not rendered correctly (theme
  colour edge case — most theme resolution is shipped via
  `resolveSchemeColor`, so this is likely a single bad dispatch).
- **#178** — tabStop display bug (decimal / bar tab styles aren't handled
  in `javascript.ts`; see below).
- **#181** — shared numbering counters across multiple previews on the
  same page (known footgun; needs scoping to the renderer root).
- **#183** — source map `sourcesContent: null` + missing src paths in
  `dist/*.map` files.
- **#187** — empty `<p>` has wrong height.
- **#195** — text-indent incorrect when `firstLineChars` is set instead
  of `firstLine`.

### Features (verified still missing)

- **`w:lang` → HTML `lang` attribute** on runs + paragraphs (currently
  only emitted on `<ruby>`). Blocks browser hyphenation and
  screen-reader language detection.
- **Decimal / bar tab stops + extra custom leaders** in
  `src/javascript.ts:58` (dot / middleDot / hyphen / heavy / underscore
  shipped).
- **`w:pBdr` art borders** (apples, stars, etc.) — `parseBorderProperties`
  only handles left/right/top/bottom.
- **`w:lvlRestart` mid-document** — parser reads the value
  (`document-parser.ts:626`) but renderer uses a CSS `counter-set` on
  level change which doesn't honour mid-level restarts.
- **Watermarks via `behindDoc`** — parsed (`document-parser.ts:1407`)
  but never applied to `z-index`, so watermark shapes render in-flow.
- **Footnote physical-bottom placement + `numRestart` policy** —
  `html-renderer.ts:334` has a TODO for threading `settings.footnoteProps`.
- **`w:settings evenAndOddHeaders` gating** — setting is parsed
  (`settings/settings.ts:29`) but `renderHeaderFooter` doesn't gate the
  even-header pick on it.
- **`tblBorders` scoping inside `tblStylePr`** — conditional sub-style
  borders over-border sibling cells.
- **Nested-table CSS inheritance isolation** — outer `tblStyle` cascades
  further than Word does.
- **Image effects on pictures**: `a:alphaModFix`, `a:duotone`, `a:lum`,
  `a:biLevel`. (Shape effects are shipped via `drawing/shapes.ts`.)
- **`a:tile` / `a:stretch` on DrawingML pictures** (VML-side tile
  shipped).
- **OLE objects** (`w:object` / `o:OLEObject`).
- **`w:oMath` `limLoc` rendering** — parsed but `<msubsup>` is always
  emitted for nary operators.
- **Legacy form fields** (`FORMTEXT` / `FORMCHECKBOX` / `FORMDROPDOWN`
  via `w:fldChar ffData`) — newer SDT-based controls are shipped.
- **Glossary documents** (`/word/glossary/document.xml`) — not loaded.
- **Data-bound SDT** from custom XML parts — custom XML parts not loaded.
- **`w:documentProtection`** — not parsed, no UI affordance.
- **`w:compat`** settings — not parsed.
- **Digital signature badge** — no surface.
- **`w:style.uiPriority` / `w:hidden`** — parsed but ignored
  (`document-parser.ts:341`).
- **SmartArt layout engine** — `<mc:Fallback>` path shipped; layout
  engine for fallback-less SmartArt remains out of scope.
- **CSP-strict mode (#12)** — no nonce threading.

### Infrastructure

- **Bound `jszip`** in `package.json` — currently `">=3.0.0"` (unbounded
  upper); bound to `^3` to prevent hypothetical 4.x breakage.
- **TypeScript `any` usages** concentrated in `html-renderer.ts` (11) and
  `document-parser.ts` (6).
- **Refresh 11 render-test golden files** listed in CHANGELOG under
  "Known".
- **Backfill CHANGELOG.md** for waves 1–8 (only 0.4.0 is documented).
- **Stale local fix branches** — `fix/comment-footnote-data-hooks`,
  `fix/continuous-footnote-numbering`, `fix/pbdr-rendering`,
  `fix/landscape-and-header-row-hooks`, `chore/overnight-n2-conformance-todo`
  (all superseded by W1-F / W9-A / W9-D / W9-E / W9-F merges).
- **Local `test-results/` + `playwright-report-interop/` clutter**
  (gitignored, worth a periodic sweep).
- **Consider a 1.0 release** that promotes `experimentalPageBreaks` off
  the experimental flag. Both upstream blockers (header/footer repeat,
  table row splits) have shipped.

---

## Methodology

- **2026-05-02** initial triage: enumerated open issues + PRs on upstream
  via `gh`; four review agents each handled ~14 items, grepping the fork's
  source tree and consulting `git log`.
- **2026-05-04** overnight corpus run against the 950-case
  `loadfix/ooxml-validate` suite produced the conformance-gaps section.
- **2026-05-07** re-audit: walked every entry in the "Missing Word
  features (read-only)", "Bugs to investigate", "Resolved in fork", and
  "Upstream PRs worth adopting" sections against master at HEAD `6479363`.
  Moved shipped items into "Resolved in fork" with PR references, removed
  the stale "Priority recommendation" block (all 5 items shipped in
  waves 1–4), trimmed the "Original entries (retained for reference)"
  subsection since the "Resolved in overnight wave 1" summary supersedes
  it, and added the "Genuinely remaining" section listing the bounded
  set of items verified as still open.

### Wave history

| Wave | PR | Focus |
|---|---|---|
| Wave 1 | #48 | Field results, wrap modes, numbering overrides, a11y, VML, bookmarks |
| Wave 2 | #49 | Paragraph renderer, char formatting, section fidelity, table fidelity |
| Wave 3 | #50 | DrawingML shapes + text boxes, equation coverage |
| Wave 4 | #51 | East-Asian/RTL, SDT controls, custom geometry, charts |
| Wave 5 | #53 | Drawing polish, chart enhancements, field-code edge cases |
| Wave 6 | #54 | Shape effects completion, SmartArt fallback, theme-wide colours |
| Wave 7 | #55 | WMF/EMF decoder, ChartEx sunburst + treemap, interop harness |
| Wave 8 | #57 | Squarified treemap, waterfall/funnel/histogram, reflection polish |
| W1-F | `d6ab27d` | 16 conformance hook fixes |
| W9-A | `071b3dd` | eastAsia/bidi w:lang fallthrough + tblLook header |
| W9-D | `832d345` | ARIA landmarks, heading roles, table semantics |
| W9-E | `fc0595f` | Opt-in responsive render option |
| W9-F | `4e39d1b` | Shared `oox-*` CSS classes |
