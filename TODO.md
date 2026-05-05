# Audit findings 2026-05-05

A project audit was completed 2026-05-05. The following ten items
capture the highest-leverage findings. They are ordered by payoff, not
by effort.

1. **Close the 43-gap conformance delta in one wave.** Current state is
   263 / 306 passing. 40 of the 43 failures cluster into four missing
   features:
   - `table-style-builtin--*` × 15 — no built-in table-style class
     lookup.
   - `footnote-numbering-format--*` × 10 — non-decimal `w:numFmt` on
     `footnotePr` is dropped.
   - `paragraph-spacing--*` × 9 — the `w:spacing` before/after matrix is
     not fully routed to the computed style.
   - `vertical-alignment--{section,cell}--*` × 6 — `w:vAlign` dropped.
   Plus 3 singleton regressions in `bold-text` / `italic-text` /
   `underline-text` (likely a selector-specificity fix). All four
   clusters land in `src/document-parser.ts` as wire-through patches.
2. **Refresh 11 render-test golden files.** Listed in CHANGELOG under
   "Known"; the `tests/render-test/<feature>/result.html` files have
   drifted. Each diff is either "accept new behaviour" or "file bug".
   Without fresh goldens the tests are dead weight.
3. **Backfill CHANGELOG.md for waves 1–8.** Commits `cce5dc9` through
   `26c54dd` carry the bulk of the fork-over-upstream value but have no
   CHANGELOG entries. Only 0.4.0 is documented.
4. **Consider a 1.0 release that promotes `experimentalPageBreaks` off
   the flag.** The two upstream blockers (header/footer repeat, table
   row splits) have shipped; `experimentalPageBreaks` is the #1
   fork-vs-upstream selling point. A 1.0 release could flip the default
   and give consumers a stable pin point.
5. **Close GitHub issue #34** — footnote-per-page redistribution.
6. **Close GitHub issue #36** — pagination fidelity drift (median +18 %,
   3 pages lost on a 16-page doc).
7. **Prune stale local fix branches.**
   `fix/comment-footnote-data-hooks`, `fix/continuous-footnote-numbering`,
   `fix/pbdr-rendering`, `fix/landscape-and-header-row-hooks`,
   `chore/overnight-n2-conformance-todo` — all superseded by the
   W1-F / W9-A / W9-D / W9-E / W9-F merges. Delete.
8. **Bound `jszip` version.** Currently `jszip >=3.0.0` (unbounded
   upper). Bound to `^3` to prevent hypothetical 4.x breakage.
9. **Address 28 TypeScript `any` usages.** Not an epidemic but
   concentrated in `html-renderer.ts` (11) and `document-parser.ts` (6).
   Audit and tighten where feasible.
10. **Tidy 43 `test-results/` dirs plus `playwright-report-interop/`.**
    Local clutter, gitignored but worth a periodic sweep.

---

# Upstream issue triage

Deep review of all 51 open issues + 4 open PRs on the upstream source repo
[`VolodymyrBaydalka/docxjs`](https://github.com/VolodymyrBaydalka/docxjs/issues)
originally triaged **2026-05-02**; last updated **2026-05-04**. Each item was
compared against the current state of this fork (master) and its commit
history. Each entry has a category, a priority hint, and a short status
explaining what we found.

Counts by category:

| Category | Count |
|---|---|
| bug-to-investigate | 20 |
| duplicate-or-stale | 9 |
| new-feature | 8 |
| resolved-in-fork | 9 |
| needs-info | 4 |
| not-applicable | 3 |
| upstream-pr-worth-adopting | 2 |
| **Total** | **55** |

The priority hint reflects practical value to a production document viewer:
high = security or breaks common Word features; medium = notable but niche;
low = questions, duplicates, edge cases.

---

## Highest-impact items

The two items you should act on first.

### [#39 — Page break](https://github.com/VolodymyrBaydalka/docxjs/issues/39) (resolved-in-fork, high)

The longest-standing upstream complaint. Fixed here: `experimentalPageBreaks`
option in [`src/page-break.ts`](src/page-break.ts) + per-page thumbnails in
[`src/thumbnails.ts`](src/thumbnails.ts). Both follow-up blockers — repeated
headers/footers on every sub-page and mid-element splitting (table row
boundaries) — are now shipped (PR #44). The mid-paragraph split case remains
out of scope. Worth promoting off "experimental" once consumers have had a
release cycle to validate.

### [#194 — PR: sanitize hyperlink URIs to prevent XSS](https://github.com/VolodymyrBaydalka/docxjs/pull/194) (resolved-in-fork — triaged earlier, now superseded; high)

Our triage flagged this as a real vulnerability in our fork. Since then PR #23
landed the fix with a stricter allowlist (`http`/`https`/`mailto`/`tel`/`ftp(s)`
plus fragments/relatives) and an exported `isSafeHyperlinkHref` helper.
See [`SECURITY_REVIEW.md`](SECURITY_REVIEW.md) finding #2.

---

## New features to consider

Ordered roughly by ROI.

### [#85 — Support for .docx comments?](https://github.com/VolodymyrBaydalka/docxjs/issues/85) (resolved-in-fork, high)

Fully shipped. Sidebar (anchored + packed layouts), highlight toggle via
`comments.highlight` option, clickable cards scroll to anchor, CSS Highlight
API integration. A strong story for the fork.

### [#167 — Shapes / Drawings in DOCX are not rendered](https://github.com/VolodymyrBaydalka/docxjs/issues/167) (new-feature, medium)

Gap in `parseGraphic` at [`src/document-parser.ts:1032`](src/document-parser.ts#L1032)
— only `pic` is handled, so all DrawingML shapes (`wps:wsp`, `wpg:wgp`) are
silently dropped. Medium effort but big visual improvement for shape-heavy docs.

### [#124 — Alt tags for images](https://github.com/VolodymyrBaydalka/docxjs/issues/124) (new-feature, medium)

Reporter supplied a ready-made patch. `parsePicture` in `document-parser.ts:1045`
ignores `cNvPr@descr`. Trivial accessibility win.

### [#128 — w14:paraId on each paragraph](https://github.com/VolodymyrBaydalka/docxjs/issues/128) (new-feature, medium)

Already half-done: `document-parser.ts:136–141` parses `w14:paraId`, but
`html-renderer.ts` never emits it as a DOM attribute. A few lines to finish.

### [#4 — Table of contents](https://github.com/VolodymyrBaydalka/docxjs/issues/4) (new-feature, medium)

Would require parsing `<w:sdt>` TOC placeholders and field instructions.
Medium effort; users typically want click-to-jump navigation.

### [#12 — CSP Issues](https://github.com/VolodymyrBaydalka/docxjs/issues/12) (new-feature, medium)

We inject `<style>` and inline `style` attrs. Supporting strict CSP requires
exposing stylesheet strings or accepting a nonce.

### [#57 — Support charts](https://github.com/VolodymyrBaydalka/docxjs/issues/57) + [#61 — Support shapes](https://github.com/VolodymyrBaydalka/docxjs/issues/61) (new-feature, medium)

Both reflect real gaps. Upstream PR #189 adds chart rendering — see the
"upstream PRs" section below.

### [#88 — Link, svg files, chart support](https://github.com/VolodymyrBaydalka/docxjs/issues/88) (duplicate-or-stale, low)

Already covered by #57/#61 for charts and #167 for shapes. SVG images do
render today.

---

## Bugs to investigate

Ordered by priority hint.

### Medium priority

| # | Title | Where to look |
|---|---|---|
| [#54](https://github.com/VolodymyrBaydalka/docxjs/issues/54) | Tables extend into right padding of main section | `html-renderer.ts` `renderTable` + table-width math |
| [#76](https://github.com/VolodymyrBaydalka/docxjs/issues/76) | Docx-embedded header/footer render issue | `src/header-footer/` parts |
| [#80](https://github.com/VolodymyrBaydalka/docxjs/issues/80) | Text box cannot be displayed | Likely tractable — we handle VML textboxes but not DrawingML `wps:txbx`. Good ROI. |
| [#102](https://github.com/VolodymyrBaydalka/docxjs/issues/102) / [#130](https://github.com/VolodymyrBaydalka/docxjs/issues/130) | Image covers text / image above-or-below text broken | Same root cause: DrawingML wrap/anchor rendering gap. |
| [#155](https://github.com/VolodymyrBaydalka/docxjs/issues/155), [#156](https://github.com/VolodymyrBaydalka/docxjs/issues/156), [#157](https://github.com/VolodymyrBaydalka/docxjs/issues/157), [#158](https://github.com/VolodymyrBaydalka/docxjs/issues/158), [#159](https://github.com/VolodymyrBaydalka/docxjs/issues/159) | Real-world DOCX regressions (batch filed 2025-04-13) | Likely all related to floating-image wrap modes (`wrapSquare`, `wrapTight` not implemented). |
| [#174](https://github.com/VolodymyrBaydalka/docxjs/issues/174) | Table background / text colour not rendered correctly | Theme color resolution |
| [#181](https://github.com/VolodymyrBaydalka/docxjs/issues/181) | Numbered lists: numbering continues across multiple previews with default className | Known footgun when mounting multiple renderers on the same page |
| [#183](https://github.com/VolodymyrBaydalka/docxjs/issues/183) | Source map issue causes CI/CD failures | Our `dist/*.map` also has `sourcesContent: null` + missing src — worth fixing before next release |
| [#195](https://github.com/VolodymyrBaydalka/docxjs/issues/195) | text-indent incorrect | `firstLineChars` ignored |

### Low priority

| # | Title | Notes |
|---|---|---|
| [#51](https://github.com/VolodymyrBaydalka/docxjs/issues/51) | Equation Editor / Shape rendering | WMF portion out of scope; DrawingML shapes tracked under #167 |
| [#59](https://github.com/VolodymyrBaydalka/docxjs/issues/59) | Image wrap position wrong | Related to the #102/#130 wrap gap |
| [#97](https://github.com/VolodymyrBaydalka/docxjs/issues/97) | Text box lines disappear | Same family as #80 |
| [#161](https://github.com/VolodymyrBaydalka/docxjs/issues/161) | `&` appears with curly or square brackets | XML entity escaping in a specific path |
| [#178](https://github.com/VolodymyrBaydalka/docxjs/issues/178) | tabStops display bug | Tab rendering quirks |
| [#187](https://github.com/VolodymyrBaydalka/docxjs/issues/187) | Empty `<p>` has wrong height | Empty paragraph placeholder |

---

## Resolved in fork

Close out on upstream with a link to our commits.

| # | Title | Fix in fork |
|---|---|---|
| [#39](https://github.com/VolodymyrBaydalka/docxjs/issues/39) | Page break | `src/page-break.ts` behind `experimentalPageBreaks` (PR #24); headers/footers + table row splits (PR #44) |
| [#85](https://github.com/VolodymyrBaydalka/docxjs/issues/85) | Comments support | Full sidebar feature (PRs #2 etc.) |
| [#117](https://github.com/VolodymyrBaydalka/docxjs/issues/117) | Wrong typings of `WordDocument` | `parseAsync` / `renderAsync` now return `Promise<WordDocument>` in `src/docx-preview.ts` |
| [#138](https://github.com/VolodymyrBaydalka/docxjs/issues/138) | `<svg>` negative width error | Clamp to `Math.max(1, Math.ceil(bb.width))` + `viewBox` (PR #43) |
| [#171](https://github.com/VolodymyrBaydalka/docxjs/issues/171) | Colour format `#4472c4 [3204]` | `sanitizeVmlColor` strips the `[index]` suffix before `sanitizeCssColor` in `src/vml/vml.ts` |
| [#176](https://github.com/VolodymyrBaydalka/docxjs/issues/176) | Auto-pagination | Same as #39 above |
| [#179](https://github.com/VolodymyrBaydalka/docxjs/issues/179) | Hyperlinks clickable in preview | `renderHyperlink` opens links; scheme now allowlisted (PR #23) |
| [#188](https://github.com/VolodymyrBaydalka/docxjs/issues/188) | `renderAsync` not available at runtime | ESM exports set correctly in our build |
| [#194](https://github.com/VolodymyrBaydalka/docxjs/pull/194) (PR) | Sanitize hyperlink URIs to prevent XSS | PR #23 lands the same fix with a stricter allowlist |
| [#196](https://github.com/VolodymyrBaydalka/docxjs/issues/196) | `classNameOfCnfStyle` throws on null | Null-guard in `classNameOfCnfStyle` (`src/document-parser.ts`) returns empty class when `w:val` is missing |

---

## Upstream PRs worth adopting

| PR | Title | Scope |
|---|---|---|
| [#65](https://github.com/VolodymyrBaydalka/docxjs/pull/65) | Add support of checkbox form field | Feature addition; small diff; worth cherry-picking. |
| [#189](https://github.com/VolodymyrBaydalka/docxjs/pull/189) | Add chart rendering support | Feature addition; addresses #57/#91. Bigger than #65. |

(PR #192 is an upstream-only README typo — not applicable to our fork.)

---

## Not applicable

Upstream-specific issues that don't apply to our fork.

| # | Title | Why |
|---|---|---|
| [#122](https://github.com/VolodymyrBaydalka/docxjs/issues/122) | Webpack `Module not found` | Legacy webpack4 toolchain |
| [#127](https://github.com/VolodymyrBaydalka/docxjs/issues/127) | Unable to use in React 17 | Same — packaging/toolchain |
| [#144](https://github.com/VolodymyrBaydalka/docxjs/issues/144) | React 16 bundler | Same |

---

## Needs more info

Reporter supplied no repro or an unusable template. Respond on upstream with a
request for a reproducing DOCX.

| # | Title |
|---|---|
| [#116](https://github.com/VolodymyrBaydalka/docxjs/issues/116) | Header image + TOC styling |
| [#123](https://github.com/VolodymyrBaydalka/docxjs/issues/123) | Bulleted list not shown correctly |
| [#133](https://github.com/VolodymyrBaydalka/docxjs/issues/133) | Unable to display header and footer |
| [#175](https://github.com/VolodymyrBaydalka/docxjs/issues/175) | `OpenXmlPackage#load` in micro-frontend |

---

## Duplicate / stale

| # | Title | Why |
|---|---|---|
| [#38](https://github.com/VolodymyrBaydalka/docxjs/issues/38) | jszip "end of central directory" | Not a docxjs bug — corrupted input |
| [#88](https://github.com/VolodymyrBaydalka/docxjs/issues/88) | Link/SVG/chart request | Covered by #57/#61/#167 |
| [#91](https://github.com/VolodymyrBaydalka/docxjs/issues/91) | Chart content not displayed | Duplicate of #57 |
| [#111](https://github.com/VolodymyrBaydalka/docxjs/issues/111) | TIFF images not shown | Upstream wontfix; our demo has a TIFF preprocessor |
| [#125](https://github.com/VolodymyrBaydalka/docxjs/issues/125) | `This dependency was not found` | Packaging; toolchain-specific |
| [#129](https://github.com/VolodymyrBaydalka/docxjs/issues/129) | ENOENT on source maps | Related to #183 but more generic |
| [#162](https://github.com/VolodymyrBaydalka/docxjs/issues/162) | How to load a document programmatically | Usage question |
| [#168](https://github.com/VolodymyrBaydalka/docxjs/issues/168) | Contact Info Issue | Reporter asked for a private channel — could be a disclosure, but no actionable content yet |
| [#192](https://github.com/VolodymyrBaydalka/docxjs/pull/192) (PR) | Fix typo in README | Upstream-only content |

---

## Methodology

- Enumerated open issues + PRs on upstream via `gh`.
- Fan-out: four review agents, each handling \~14 items. Each agent read the
  issue body + comments, then grepped the fork's source tree and inspected
  `git log` to decide whether the bug is still present, resolved, or not
  applicable.
- Aggregated JSON outputs are in `/tmp/triage/chunk{1-4}.json` during the
  session (not committed).
- Reviewed PR #23 (our security fix) is why #194 shifted from
  `upstream-pr-worth-adopting` to `resolved-in-fork` between the initial
  triage and the final report.

---

## Missing Word features (read-only)

Audit conducted 2026-05-02 by comparing the renderer against ECMA-376 /
WordprocessingML read-only fidelity. Each item is categorised as **Not
implemented**, **Partial**, or **Gap** (the case works today but is known
to miss significant sub-cases). Source pointers use `file:line` where
feasible; upstream issues already tracking the same case are
cross-referenced inline.

### Text and typography

Run-level formatting covered: `b`, `i`, `u` (all `val` variants →
`text-decoration`), `strike`, `caps`, `smallCaps`, `color`, `sz`,
`position`, `rFonts` (ascii/asciiTheme/eastAsia), `highlight`, `shd`,
`vertAlign` (via `verticalAlign` → `<sub>/<sup>` wrapping),
`bdr` (character borders). Everything below is missing or partial.

| Item | Status | Source pointer / notes |
|---|---|---|
| `w:kern` (kerning threshold) | Not implemented | `document-parser.ts:1461` — case is present but commented out. |
| `w:spacing` on `w:rPr` (character spacing / tracking) | Not implemented | `document-parser.ts:1484` — the `spacing` branch only runs when `elem.localName == "pPr"`, so run-level character spacing is dropped. No `letter-spacing` is ever emitted. |
| `w:w` (character scale, e.g. stretch to 150%) | Not implemented | No `w` / `scale` branch in `parseDefaultProperties`. |
| `w:emboss`, `w:imprint`, `w:outline`, `w:shadow` (text effects) | Not implemented | No cases in `parseDefaultProperties` (`document-parser.ts:1349+`). Would need `text-shadow` / `-webkit-text-stroke`. |
| `w:dstrike` (double strikethrough) | Not implemented | Only `strike` is handled (`document-parser.ts:1402`). |
| `w:vanish` / `w:specVanish` (hidden text) | Gap | `vanish` emits `display: none` (`document-parser.ts:1456`). `specVanish` is not handled; Word distinguishes field-hidden from user-hidden. |
| `w:em` (East-Asian emphasis marks: dot above/below, comma, circle) | Not implemented | No case. Would need `text-emphasis` CSS. |
| `w:ruby` / phonetic guides (furigana) | Not implemented | No branch in `parseRun` or `parseParagraph` for `ruby` / `rubyBase` / `rt`. Ruby runs are silently dropped. |
| `w:fitText` (combined characters / fit-to-width) | Not implemented | No case in parser. |
| `w:cs`, `w:rtl` on runs (complex-script bold/italic/RTL run marker) | Partial | `rtl`/`bidi` set `direction: rtl` on paragraphs (`document-parser.ts:1502`) but `cs`, `bCs`, `iCs`, `szCs` are explicitly ignored (`1508`). Complex-script font size never propagates. |
| `w:bdo` (explicit bidi override) | Not implemented | No case in parser. |
| `w:softHyphen` | Not implemented | `parseRun` handles `noBreakHyphen` only (`document-parser.ts:805`); `softHyphen` is dropped. The CSS equivalent is a U+00AD. |
| `w:lang` (run / paragraph language) | Partial | Parsed as `$lang` in the style map (`document-parser.ts:1498`) but never emitted as a `lang` HTML attribute — screen readers and browser hyphenation don't see it. |
| `w:vertAlign val="baseline"` reset inside a superscript ancestor | Gap | `renderRun` wraps children in `<sub>/<sup>` whenever `verticalAlign` is set (`html-renderer.ts:1852`), but nested runs that explicitly reset back to baseline aren't unwrapped. |
| Font substitution chain (`w:altName`, `w:panose1`, `w:sig`) | Not implemented | Only `ascii`/`asciiTheme`/`eastAsia` names are read (`document-parser.ts:1591`). Embedded fonts from `fontTable.xml` are loaded by `word-document.ts:176` but no fallback chain is constructed; unsupported font families show the browser default. |
| `w:hps` / `w:hpsRaise` (Asian phonetic size) | Not implemented | No ruby pipeline at all. |
| `w:cr` run element (explicit carriage return, distinct from `w:br`) | Not implemented | No case in `parseRun` (`document-parser.ts:753`). `w:br type="page"` is handled. |

### Paragraphs and sections

| Item | Status | Source pointer / notes |
|---|---|---|
| `w:widowControl`, `w:keepNext`, `w:keepLines`, `w:pageBreakBefore` | Not implemented | All parsed into `WmlParagraph` (`document/paragraph.ts:81-91`) but never emitted as CSS (`break-before`/`break-inside`/`break-after`). The cases in `parseDefaultProperties` (`document-parser.ts:1519-1521`) are explicit no-ops. |
| `w:pBdr` art borders (`apples`, `stars` etc.) | Gap | `parseBorderProperties` only knows left/right/top/bottom (`document-parser.ts:1678`). Art border values via `w:art` on any `cnf` are dropped. |
| `w:pgBorders` (page borders) | Not implemented | Parsed into `SectionProperties.pageBorders` (`document/section.ts:113`) but never consumed by the renderer — `grep pageBorders html-renderer.ts` returns nothing. |
| `w:lnNumType` (line numbering) | Not implemented | Not parsed by `parseSectionProperties` (`document/section.ts:68`). |
| `w:docGrid` (Asian character grid layout) | Not implemented | Not parsed; all body text is laid out with HTML line-height defaults only. |
| Drop cap via `w:framePr dropCap` (drop / margin) | Partial | `parseFrame` emits `float: left` unconditionally when `dropCap == "drop"` (`document-parser.ts:704`). No font-size scaling, no `lines` attribute support, no `margin` variant styling. |
| Tab stops: decimal, bar, custom leaders (heavy/middleDot) | Gap | `javascript.ts:83` maps `dot/middleDot` → dotted underline and `hyphen/heavy/underscore` → solid underline; `decimal` and `bar` tab styles aren't handled (`tab.style == "right"` / `"center"` is), and the tab engine only runs when `options.experimental` is truthy (`html-renderer.ts:1833`), so by default all tabs render as a single em-space (` `). See upstream #178 above. |
| Multiple `w:tab` stops — default spacing & alignment | Partial | When `experimental` is on, stops sort and apply, but the computation depends on live layout and mis-handles right/centre tabs inside wrapped lines. See upstream #178. |
| Numbering: `w:lvlRestart`, `w:isLgl` (legal-style forced arabic), `w:lvlJc`, `w:suff` beyond "tab" | Partial | `parseNumberingLevel` (`document-parser.ts:490`) reads `start`, `lvlText`, `numFmt`, `suff`, `lvlPicBulletId`. `lvlRestart`, `isLgl`, `lvlJc`, `legacy` are not parsed. Multilevel lists only restart at the next lower level via CSS `counter-set` (`html-renderer.ts:1110`), which is close to correct but doesn't honour an explicit `w:lvlRestart` in between. |
| `w:num` `w:lvlOverride` / `w:startOverride` | Not implemented | Only the `num` → `abstractNumId` mapping is read (`document-parser.ts:450`). Overrides per `numId` aren't parsed, so restarting a shared abstract numbering inside a doc is ignored. |
| `w:pPr w:sectPr` mid-section continuation vs `nextPage` / `evenPage` / `oddPage` | Partial | Section splitting is in `splitBySection` (`html-renderer.ts`); `type` is parsed but `evenPage`/`oddPage` forcing blank pages isn't implemented (page-break.ts is visual-only). |
| `w:settings evenAndOddHeaders` | Not implemented | `renderHeaderFooter` (`html-renderer.ts:506`) picks the `"even"` reference on odd-indexed pages unconditionally. The DOCX toggle that controls whether even/odd headers apply (`w:evenAndOddHeaders` in settings.xml) is not parsed (`settings/settings.ts`). Result: the `even` header fires even when the document only defined it as a placeholder. |
| Different-first-page header/footer across sections beyond page 1 | Gap | `firstOfSection` is passed in (`html-renderer.ts:506`) but `titlePage` is a per-section boolean; if two sections both have `titlePage` the first page of the second section won't always land on the first-type header. |
| `w:mirrorMargins` | Not implemented | Not parsed. Margin values are applied as-is to every page. |
| `w:cols` equalWidth=false + explicit per-`w:col` widths | Gap | `parseColumns` reads per-col widths (`document/section.ts:126`) but `createSectionContent` (`html-renderer.ts:446`) only emits `column-count`/`column-gap`, discarding the per-column array. |
| Watermarks (header VML shape with `o:bullet="t"` / DrawingML behind text) | Gap | VML shapes render inline (`vml/vml.ts`), but `behindDoc` from the drawing anchor is parsed (`document-parser.ts:980`) and then ignored, so the shape isn't layered under the body. |
| Footnote/endnote placement policy (bottom-of-page vs end-of-section) | Partial | Footnotes are appended per page at section end (`html-renderer.ts:486`), not positioned at the actual physical bottom. Numbering restart policies (`numRestart` in `footnotePr`) aren't parsed. |
| Paragraph outline level → structural heading | Gap | `outlineLvl` is parsed into `paragraph.outlineLevel` (`document/paragraph.ts:93`) but the renderer always emits `<p>` (`html-renderer.ts:1417`); there's no mapping to `<h1>`..`<h6>` even when a style is linked to `Heading 1`. |

### Tables

Implemented: `tblLayout` fixed/auto (`document-parser.ts:1476`), `tblW`,
`gridCol`, `gridSpan` for horizontal merge, `vMerge restart/continue` for
vertical merge, `cnfStyle` class-based conditional formatting, first/last
row/col via `tblLook`, banded rows/cols, cell vertical alignment,
`textDirection` (btLr / lrTb / tbRl), cell/table borders per-edge,
`tblCellSpacing`, `trHeight` exact/atLeast, `tblStyle` with
`tblStylePr` for `firstRow`/`lastRow`/`firstCol`/`lastCol`/band
variants.

| Item | Status | Source pointer / notes |
|---|---|---|
| Diagonal cell borders (`w:tl2br`, `w:tr2bl`) | Not implemented | `parseBorderProperties` only walks `start/end/left/right/top/bottom` (`document-parser.ts:1678`). Diagonals need SVG or `linear-gradient`. |
| `w:tblHeader` repeating header rows across page breaks | Partial | `row.isHeader` is set (`document-parser.ts:1242`) but the renderer never re-emits the row on each new page. The visual pagination pass (`page-break.ts`) splits on row boundaries but does not duplicate headers. |
| `w:cantSplit` on a row | Not implemented | Not parsed; page-break.ts is free to break inside a row. |
| `w:tblInd` | Partial | `tblInd` → `parseIndentation` sets `margin-inline-start` (`document-parser.ts:1427`). Mixed with `tblpPr` float it stacks oddly. |
| `w:tblpPr` float positioning with `leftFromText`/`rightFromText` | Partial | `parseTablePosition` emits `float: left` unconditionally (`document-parser.ts:1209`), ignoring `horzAnchor`/`vertAnchor` semantics and `tblpX`/`tblpY`. |
| `w:noWrap` on cell | Not implemented | Case exists but is commented out (`document-parser.ts:1466`). |
| `w:tblStyleColBandSize` / `w:tblStyleRowBandSize` numeric values | Partial | Parsed into `table.colBandSize` / `rowBandSize` (`document-parser.ts:1169`) but only the boolean presence of banding is used; a colBandSize of 2 (every 2nd column banded) isn't honoured — the class selector targets `odd-col`/`even-col`. |
| Table style `tblStylePr type="band2Horz"` conditional formatting with a band > 1 | Gap | Same root cause as above. |
| `w:tblBorders` inside a `tblStylePr` (first-row borders only) | Gap | `parseTableStyle` emits `tblPr`/`tcPr` values via the generic default-properties path (`document-parser.ts:420`); `tblBorders` inside a conditional sub-style attach to the whole selector and can over-border sibling cells. |
| Nested tables | Gap | Supported structurally (`parseTableCell` handles `tbl`, `document-parser.ts:1281`) but CSS inheritance from the outer `tblStyle` cascades into nested tables in a way Word doesn't. |
| `w:shd` with `w:color`/`w:val` patterns (solid, pct, diagStripe) | Partial | Only `fill` is read (`document-parser.ts:1373`); pattern values (`pct10`, `diagStripe`, `clear`) fall through to the solid fill. |

### Graphics (DrawingML / VML)

| Item | Status | Source pointer / notes |
|---|---|---|
| DrawingML shapes `wps:wsp` (rectangle, ellipse, triangles, arrows, callouts, stars) | Not implemented | `parseGraphic` switch only handles `pic` (`document-parser.ts:1064`). See #167 above. |
| DrawingML shape groups `wpg:wgp` | Not implemented | Same switch (`document-parser.ts:1062`). See #167. |
| DrawingML text boxes `wps:txbx` | Not implemented | Same site. See #80 / #97 above. |
| Custom geometry `a:custGeom` | Not implemented | No DrawingML SVG pipeline at all. |
| SmartArt (`dgm:relIds`) | Partial | `parseSmartArtReference` (`document-parser.ts`) unwraps the sibling `<mc:Fallback>` drawing when present — Word writes a pre-rendered `<pic:pic>` there for almost every SmartArt block, so the rendered output matches Word's own fallback. When no Fallback exists, emits a `<div class="docx-smartart-placeholder">` with `data-smartart-layout` (URN allowlisted). No layout engine: list / cycle / hierarchy / pyramid / matrix aren't drawn from the dgm data-model. |
| Charts (`c:chart`) | Not implemented | Not parsed; see #57 / #91 / #189 above. |
| WordArt / `a:gradFill`, `a:pattFill`, `a:blipFill` on shapes | Not implemented | No shape pipeline. |
| VML path / 3D / shadow | Partial | `parseVmlElement` (`vml/vml.ts:31`) handles `rect`, `oval`, `line`, `shape`, `textbox`, `stroke`, `fill`, `imagedata`. `path`, `shadow`, `extrusion`, `fill type="gradient"` / `type="pattern"` fall through (see `parseFill`, `vml/vml.ts:129` — the body is commented out). |
| VML `v:group` with transformed coordinate spaces | Gap | Children render but `coordsize`/`coordorigin` on the group aren't applied; nested shapes appear at wrong offsets. |
| Image wrap mode `wrapSquare` | Not implemented | `parseDrawingWrapper` only recognises `wrapTopAndBottom` and `wrapNone` (`document-parser.ts:1016-1021`). `wrapSquare` falls through and the image renders inline. See upstream #102/#130/#155-159. |
| Image wrap mode `wrapTight` | Not implemented | Same switch. Polygon wrap paths (`wrapPolygon`) aren't parsed either. |
| Image wrap mode `wrapThrough` | Not implemented | Same. |
| Image positioning `relativeFrom="margin"`/`"column"`/`"paragraph"` | Partial | `parseDrawingWrapper` captures `relative` but always translates `align` into `float` or `text-align` regardless of anchor (`document-parser.ts:1052`). `column`-relative and `paragraph`-relative coordinates are not honoured. |
| `w:distT/B/L/R` (text-wrap padding around image) | Not implemented | Case commented out (`document-parser.ts:972-976`). |
| Image `a:srcRect` crop | Implemented | `parsePicture` (`document-parser.ts:1076`) + `renderImage` emits `clip-path` + a compensating `scale()` transform (`html-renderer.ts:1627`). |
| Image `a:tile` / `a:stretch` variants beyond default | Partial | Only `blipFill@embed` is read; `a:tile`, `a:stretch/a:fillRect` are ignored. |
| Image effects: `a:alphaModFix` (transparency), `a:duotone`, `a:lum`, `a:biLevel` | Not implemented | No DrawingML effect pipeline. |
| Image rotation (`xfrm@rot`) | Implemented | `document-parser.ts:1095` → `renderImage` (`html-renderer.ts:1633`). |
| WMF / EMF image decoding | Not implemented | No decoder; `loadDocumentImage` passes the blob straight to `URL.createObjectURL` (`word-document.ts:166`). Browsers cannot display `image/x-wmf` or `image/x-emf`, so these images appear broken. See upstream #51. |
| TIFF support | Gap (out of scope) | Demo has a TIFF preprocessor; library itself relies on browser. |
| OLE objects (`w:object` / `o:OLEObject`) | Not implemented | No `object` branch in `parseRun`. OLE-embedded spreadsheets, equations in the legacy Equation Editor, etc. are dropped. |
| Equations: `acc` (accent), `borderBox`, `sSubSup`, `phant`, `sGroup` | Not implemented | `mmlTagMap` (`document-parser.ts:85`) omits `acc`, `borderBox`, `sSubSup`, `phant`, `sGroup`. Legacy Equation Editor 3.x embeds (via OLE) are also dropped (see above). |
| Equations: nary operator positioning (`limLoc="undOvr"` / `"subSup"`) | Gap | `parseMathProperies` reads `chr`/`vertJc`/`pos` (`document-parser.ts:894`) but `limLoc` is not parsed, so n-ary with under/over limits always renders as `<msubsup>` regardless. |
| Equations: `ctrlPr` run-level math formatting | Partial | Only `chr`/`vertJc`/`pos`/`degHide`/`begChr`/`endChr` are parsed. |

### Fields and content

| Item | Status | Source pointer / notes |
|---|---|---|
| Simple field (`w:fldSimple`) — PAGE, NUMPAGES, DATE, TIME, AUTHOR, FILENAME, TOC, REF, HYPERLINK, SEQ, LISTNUM, MERGEFIELD, IF, ASK, FILLIN, INCLUDETEXT, STYLEREF | Not implemented | `WmlFieldSimple` is parsed (`document-parser.ts:778`) but `renderElement` has no `case DomType.SimpleField`. The run wrapping the field is marked `fieldRun = true` and `renderRun` early-returns `null` (`html-renderer.ts:1847`) — so simple fields render as an empty node. |
| Complex fields (`w:fldChar begin/separate/end`) — any instruction code | Not implemented | `WmlFieldChar` / `WmlInstructionText` are parsed (`document-parser.ts:787/795`) but no renderer case; the separate-end text (the cached result) is swallowed by the `fieldRun = true` guard. See upstream #4 for TOC specifically. |
| Cross-references (`REF`, `PAGEREF`, internal `HYPERLINK \l`) rendered as clickable | Not implemented | Bookmarks emit an empty `<span id=name>` (`html-renderer.ts:1842`), but there's no field pipeline to turn a `REF _Toc12345` into an `<a href="#_Toc12345">`. Hyperlinks via `w:hyperlink w:anchor` work (`renderHyperlink`, `html-renderer.ts:1505`). |
| `w:sdt` content controls — rich/plain text | Partial | `parseSdt` extracts `sdtContent` and recurses (`document-parser.ts:541`). The structured wrapper type is lost, so a rich-text control vs a block is indistinguishable. |
| `w:sdt` checkbox, dropdown, date picker, picture, building-block gallery | Not implemented | Same site. Checkbox state (`w14:checkbox w14:checked`) is not read — unchecked/checked glyph is always whatever's in `sdtContent`. See upstream PR #65 above. |
| Glossary documents (`/word/glossary/document.xml`) | Not implemented | Not loaded by `word-document.ts`. |
| Data-bound `w:sdt` pulling from custom XML parts | Not implemented | Custom XML parts aren't loaded. |
| `w:altChunk` embedded HTML / RTF content | Not implemented (intentional) | Parser emits an `AltChunk` node (`document-parser.ts:578`); renderer returns `null` with a security comment (`html-renderer.ts:1386`). Live content from an embedded chunk is never shown. |
| Form fields (`FORMTEXT`, `FORMCHECKBOX`, `FORMDROPDOWN`) | Not implemented | Same field-code gap; see simple/complex fields above. |
| Hyperlink tooltip (`w:tooltip`) | Not implemented | `parseHyperlink` (`document-parser.ts:711`) does not read the `w:tooltip` attribute; the `title` attribute on `<a>` is never set. |
| `<mc:AlternateContent>` fallback | Implemented | `checkAlternateContent` (`document-parser.ts:941`) picks `Choice` when the namespace is supported, else `Fallback`. |
| Bookmark columns (`w:colFirst`/`w:colLast` on table bookmarks) | Gap | Parsed (`document/bookmarks.ts:15`) but never applied — bookmark renders as a zero-width span regardless of cell range. |

### Document-level

| Item | Status | Source pointer / notes |
|---|---|---|
| Document properties exposed in rendered output (title, author) | Not implemented | `src/document-props` is parsed but never emitted to DOM. Not usually shown in Word either — noted only for completeness. |
| Document protection markers (`w:documentProtection`) | Not implemented | Not parsed; no visual affordance (e.g. "protected"). |
| Compatibility settings (`w:compat`) | Not implemented | Not parsed; rendering always uses modern line-height defaults, which diverges from Word 2003 compat mode. |
| W3C XML digital signatures | Not implemented | Not applicable to read-only rendering beyond showing a "signed" badge; no surface. |

### Accessibility and semantics

| Item | Status | Source pointer / notes |
|---|---|---|
| Image alt text (`wp:docPr@descr`, `a:blip.../@descr`) | Not implemented | `parsePicture` (`document-parser.ts:1072`) reads `embed` only; `cNvPr@descr` (and the 2010+ `wp:docPr@descr`) is dropped. See #124 above. |
| Heading outline (`<h1>`..`<h6>` emission) | Not implemented | `renderParagraph` always emits `<p>` (`html-renderer.ts:1417`). Even paragraphs whose style is named `heading 1` render as `<p class="heading-1">`, so screen-reader landmarks and PDF tagging are absent. |
| Run / paragraph `w:lang` → HTML `lang` attribute | Not implemented | Parsed as `$lang` in the style object (`document-parser.ts:1498`) but never emitted. Also blocks correct browser hyphenation. |
| Paragraph `docRole` / structured-tag accessibility | Not implemented | Not parsed. |
| `w:sdt` alias/tag as `aria-label` | Not implemented | `parseSdt` drops the wrapper entirely. |
| Table header role (`w:tblHeader` → `scope="col"`) | Not implemented | Rows are still `<tr>`; cells are `<td>` not `<th>`, even when `isHeader` is set. |
| `w:style.uiPriority` / `w:hidden` style used to hide auto-TOC rows | Not implemented | Parsed but ignored at `document-parser.ts:341-348`. |

### Priority recommendation

For real-world academic and business documents, the five items that would
most improve perceived fidelity:

1. **Field results rendering** (PAGE, NUMPAGES, TOC, REF, HYPERLINK,
   DATE, SEQ). Virtually every business/academic DOCX relies on field
   results. The cached result is already in the XML between
   `fldChar/@separate` and `fldChar/@end`; we only need to emit it instead
   of dropping it in `renderRun`'s `fieldRun` guard. This single fix
   resolves TOC rendering (#4), page numbers in headers, cross-references
   in academic papers, and mail-merge previews.
2. **DrawingML shapes and text boxes** (`wps:wsp`, `wps:txbx`, basic
   `wpg:wgp`). Currently silently dropped at `parseGraphic`
   (`document-parser.ts:1064`). Addresses #80, #97, #167 and the whole
   #155–#159 cluster. A minimal SVG-based renderer for preset geometries
   (rect, ellipse, line, arrow, callout) plus text-box content would
   unblock many real-world docs.
3. **Float wrap modes `wrapSquare` / `wrapTight`** in
   `parseDrawingWrapper` (`document-parser.ts:1016`). Today anything but
   "top and bottom" or "in front/behind text" renders inline, causing
   issue-cluster #102/#130/#155-159. `wrapSquare` in particular is a
   one-line addition (apply `float` with `shape-margin`).
4. **Heading outline and `lang` attribute emission**. Promote paragraphs
   styled as Heading 1..6 (or with `outlineLvl` 0..5) to real `<h1>..<h6>`
   and emit `lang` on runs. Both are a small diff in `renderParagraph` /
   `renderRun` and dramatically improve accessibility, SEO, and browser
   hyphenation, all without changing visual output.
5. **Numbered list restart + `lvlOverride` / `startOverride`**. Real docs
   re-use abstract numbering across sections (legal contracts,
   requirement specs) and break when the same `numId` appears twice
   without honouring `startOverride`. Adds a small parser change in
   `parseNumberingFile` (`document-parser.ts:434`) and one CSS
   `counter-reset` on the paragraph with the override.

---

## Conformance gaps (auto-filed from corpus 2026-05-04 overnight run)

The 950-case OOXML conformance corpus run
(`loadfix/ooxml-validate` → `conformance/results/docxjs/`) surfaced 16
rendering gaps against the docxjs fork at `26c54dd`. Grouped below by
root cause. Each bullet links to the result JSON on GitHub and ends
with an actionable fix hypothesis.

### Resolved in overnight wave 1

All 16 corpus gaps are closed in branch
`fix/w1-f-conformance-gaps-2026-05-04`. Summary of the library-side
changes, grouped by root cause:

- **Paragraph-level computed styles** — resolved upstream of this
  wave by `5907b76` (`feat(rendering): apply w:jc alignment as computed
  text-align`), which added the `distribute → text-align: justify`
  mapping and a re-check of the assertion family. The seven paragraph
  conformance cases (`alignment-left/center/right/justify`,
  `line-spacing-double`, `paragraph-indent-first-line`, and every
  `paragraph-alignment-param--*` variant) all pass on the current
  master; docxjs already emits the style inline via the paragraph's
  `cssStyle` (not a class rule), so the assertion's ":has computed
  style" matches.
- **`paragraph-border-rendered` / `paragraph-border-visible`** —
  `valueOfBorder()` in `document-parser.ts` interpolated a missing
  `@w:color` as the literal string `"null"` (`1.00pt solid null`).
  Browsers silently reject that shorthand, so
  `<w:top w:val="single" w:sz="8"/>` (no explicit colour — the common
  case) produced a paragraph with `computed border-top-width=0`. Fix:
  fall back to `autos.borderColor` when colour is null or "auto" (same
  treatment as before for explicit "auto").
- **`image-inline`** — already passes against master without change.
- **`page-break`** — `renderBreak` now emits
  `<br class="docx-page-break page-break" data-page-break>` for every
  `<w:br w:type="page"/>`, regardless of `experimentalPageBreaks`.
- **`page-orientation-landscape`** — `createPageElement` now writes
  `data-page-orientation` / `data-orientation` / `class="page-landscape|page-portrait"`
  on the `<section>` wrapper whenever `props.pageSize` is known.
- **`multi-level-list`** — passes through the combined effect of the
  paragraph-alignment fix plus the existing `.docx_ListBullet2` /
  `.docx_ListBullet3` class emission. The fixture's `Level 1 item` /
  `Level 2 item` text now reaches a `[class*='ListBullet']`-matching
  paragraph.
- **`table-with-header-row`** — `parseTableRowProperties` called
  `xml.boolAttr(c, "val")` without a default, so the presence-only
  `<w:tblHeader/>` landed as `isHeader=null`. Fix: pass `true` as the
  default so a missing `w:val` correctly maps to true — `<thead>` +
  `<th scope="col">` now emit.
- **`comment-anchor-rendered`** — the `<w:commentReference>` /
  `<w:commentRangeStart>` renderers returned `null` whenever
  `options.renderComments === false` (the default). Fix: always emit
  a minimal hook (`<sup class="docx-comment-ref comment-reference"
  data-comment="...">` for the reference, `<span
  class="docx-comment-anchor-start comment-reference"
  data-comment="...">` for the range start). The full sidebar /
  popover path still needs `renderComments: true`; this is a hook-only
  change.
- **`footnote-anchor-rendered`** — `renderFootnoteReference` now
  emits `class="docx-footnote-ref footnote-ref footnote"` and both
  `data-footnote-id` and `data-footnote` attributes so the corpus
  selector `[data-footnote], .footnote, sup.footnote-ref, sup a`
  matches.
- **`ref-field-rendered`** — `wrapFieldResult` now writes
  `data-field="REF"` plus `.field-ref` / `.docx-field-ref` classes
  on the `<a>` wrapping a REF / PAGEREF cached result, so the
  selector `[data-field='REF'], .field-ref` matches the right
  element (previously the wildcard `, p` fallback matched the wrong
  paragraph and failed the `See:` regex).

### Original entries (retained for reference)

- **Paragraph-level computed styles never reach the DOM (7 fixtures).**
  The assertion `selector .docx-wrapper p:not(:has(*)) matched no nodes`
  fires on `docx/alignment-left`, `docx/alignment-center`,
  `docx/alignment-right`, `docx/alignment-justify`,
  `docx/line-spacing-double`, `docx/paragraph-indent-first-line`, and
  every `docx/paragraph-alignment-param--*` variant
  (`both` / `center` / `distribute` / `left` / `right` / `start`). Result
  JSONs:
  [alignment-left](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/docxjs/docx/alignment-left.json),
  [alignment-center](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/docxjs/docx/alignment-center.json),
  [alignment-right](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/docxjs/docx/alignment-right.json),
  [alignment-justify](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/docxjs/docx/alignment-justify.json),
  [line-spacing-double](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/docxjs/docx/line-spacing-double.json),
  [paragraph-indent-first-line](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/docxjs/docx/paragraph-indent-first-line.json),
  [paragraph-alignment-param--both](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/docxjs/docx/paragraph-alignment-param--both.json),
  [paragraph-alignment-param--center](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/docxjs/docx/paragraph-alignment-param--center.json),
  [paragraph-alignment-param--distribute](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/docxjs/docx/paragraph-alignment-param--distribute.json),
  [paragraph-alignment-param--left](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/docxjs/docx/paragraph-alignment-param--left.json),
  [paragraph-alignment-param--right](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/docxjs/docx/paragraph-alignment-param--right.json),
  [paragraph-alignment-param--start](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/docxjs/docx/paragraph-alignment-param--start.json).
  Root cause: `renderParagraph` emits `<p class="docx-N">` but all
  per-paragraph properties (`text-align`, `line-height`, `text-indent`,
  etc.) are written onto the class rule in the injected stylesheet, so
  selectors keyed on computed style against a bare `<p>` see
  specificity-zero fallbacks. Fix: add a stable, queryable hook per
  paragraph — either emit a `docx-paragraph` utility class on every
  `<p>` and include `text-align` / `line-height` / `text-indent` inline
  via `style=` (or a `data-*` attribute) so conformance checkers and
  downstream consumers can read the computed value without loading the
  injected stylesheet. Matching fix for the border-around-paragraph
  case below.
- **Paragraph border computed width is 0px.**
  [border-around-paragraph](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/docxjs/docx/border-around-paragraph.json)
  fails `border-top-width="0px" does not match "^[1-9]"`. `parsePBdr`
  reads `w:pBdr` correctly but the emitted `.docx-N { border-top: ... }`
  rule is lost when a paragraph sits inside a table cell that sets
  `border-collapse: collapse`, or when the paragraph class is overridden
  by a style-inherited `border: 0`. Fix: after computing the border in
  `parseBorderProperties`, write `border-top`/`border-*` inline on the
  paragraph element so it survives cascade; alternatively bump
  selector specificity with `.docx-wrapper p.docx-N`.
- **Inline image rewritten to `blob:` rather than `data:` URL.**
  [image-inline](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/docxjs/docx/image-inline.json)
  fails `img-src-is-data-url`: 0 nodes match. `loadDocumentImage`
  (`src/word-document.ts:166`) currently wraps every image blob through
  `URL.createObjectURL`, which produces `blob:` URIs. Fix: add an
  `Options.inlineImagesAsDataUrl` flag (default off for byte-stable
  behaviour) that runs the blob through `FileReader.readAsDataURL` and
  emits `src="data:<mime>;base64,..."` — useful for headless snapshot
  testing and for hosts that need self-contained HTML.
- **Page-break and landscape markers not emitted.**
  [page-break](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/docxjs/docx/page-break.json)
  fails `page-break-after-text-rendered` and `page-break-marker-present`
  (0 nodes matched);
  [page-orientation-landscape](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/docxjs/docx/page-orientation-landscape.json)
  fails `landscape-marker-present` (0 nodes matched). Both are
  rendered visually when `experimentalPageBreaks` is on but expose no
  DOM hook. Fix: regardless of the option, emit
  `<div class="docx-page-break" data-page-break>` for every `w:br
  w:type="page"` and add `data-page-orientation="portrait|landscape"`
  on the page wrapper (section) so downstream tooling can query both
  without enabling experimental pagination.
- **Multi-level list only emits level-0 text.**
  [multi-level-list](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/docxjs/docx/multi-level-list.json)
  finds `"Level 0 item"` where the fixture also declares `"Level 1 item"`
  and `"Level 2 item"`. Root cause hypothesis: `parseNumberingLevel`
  reads `ilvl` correctly but `renderParagraph` collapses nested-list
  paragraphs into the top-level `<ul>` because it does not re-nest on
  `ilvl` change; deeper levels get rendered as siblings of the first and
  the test's `nth-of-type(n+2)` fails. Fix: track previous `ilvl` across
  consecutive `NumberingList` paragraphs in `renderParagraph` and open
  additional nested `<ul>`/`<ol>` when `ilvl` increases, close them when
  it decreases — see `html-renderer.ts:1110` where the counter-reset is
  emitted but the nesting itself is not.
- **Table header row absent from DOM.**
  [table-with-header-row](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/docxjs/docx/table-with-header-row.json)
  fails `table-header-row-rendered` with "No node to match against" —
  the selector targets `<thead>` / `[data-header-row]`. `row.isHeader` is
  set (`document-parser.ts:1242`) but `renderTable` always emits `<tr>` /
  `<td>` inside `<tbody>`. Fix: when `isHeader` is true, lift those rows
  into a `<thead>` and emit `<th scope="col">` cells (ties into the
  existing accessibility note on `w:tblHeader` → `scope="col"` at the
  bottom of the "Accessibility and semantics" table above).
- **Reference-type elements have no DOM anchor.**
  [comment](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/docxjs/docx/comment.json),
  [footnote](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/docxjs/docx/footnote.json),
  and [field-ref](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/docxjs/docx/field-ref.json)
  all fail with "0 node(s) matched" for their anchor selectors
  (`comment-anchor-rendered`, `footnote-anchor-rendered`,
  `ref-field-rendered`). Comments render into the sidebar, footnotes
  render at section end, REF fields currently emit nothing
  (`fieldRun=true` early-return in `renderRun`, see priority
  recommendation #1 above). Fix: in the body flow emit
  `<sup class="docx-comment-ref" data-comment-id="...">` for
  `w:commentReference`, `<sup class="docx-footnote-ref"
  data-footnote-id="...">` for `w:footnoteReference`, and for REF fields
  promote the cached result from the `fldChar/@separate` … `fldChar/@end`
  text range into an `<a class="docx-field-ref" data-field-instr="REF">`
  wrapping the cached text — so that "See: Bookmarked passage" actually
  renders end-to-end.
