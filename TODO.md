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
