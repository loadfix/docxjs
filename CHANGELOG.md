# Changelog

## 0.4.0 — 2026-05-05

Conformance and accessibility push. Five branches consolidated into one
release: three conformance-fix branches (W1-F and W9-A pass 263/306
conformance fixtures) and two feature branches (W9-D a11y, W9-E
responsive, W9-F shared `oox-*` CSS classes).

### Added

- **Cross-format shared CSS classes (W9-F).** Every rendered structural
  element now carries an `oox-*` class alongside its `docx-*` class —
  `oox-page`, `oox-wrapper`, `oox-paragraph`, `oox-heading`, `oox-table`,
  `oox-image`. Manifests and consumer stylesheets can select "a page of
  any OOXML format" uniformly across docxjs / pptxjs / xlsxjs. The
  existing `.docx-*` classes are unchanged — this is strictly additive.
- **Opt-in responsive rendering (W9-E).** New `responsive: true` option
  on `renderAsync` / `renderDocument` replaces the fixed pixel-based
  page frame with a fluid width container so documents reflow at small
  viewport sizes. Off by default.
- **Accessibility landmarks and semantics (W9-D).**
  - `role="document"` on the docx-wrapper so screen readers treat the
    body as a single coherent reading region.
  - `scope="col"` / `scope="colgroup"` on first-row cells when
    `tblLook` marks the row as styled (Word's common "header via
    visual styling" case). Explicit `<w:tblHeader/>` rows were already
    routed to `<thead>` with `<th scope="col">` — this covers the
    implicit case.
  - Full per-script `w:lang` fallthrough: `@w:val → @w:eastAsia →
    @w:bidi` so a Japanese-only or RTL-only run still surfaces a
    BCP-47 language hint for assistive tech and hyphenation.

### Fixed

- **16 conformance hooks (W1-F).** Queryable `data-*` attributes,
  page-orientation hooks on `<section>`, a `data-page-break` marker on
  manual page breaks, `data-footnote`/`data-footnote-id` on footnote
  refs, and a border-color fallback that no longer interpolates `null`
  into the CSS shorthand (closing the `border-around-paragraph`
  fixture failure).
- **Conformance configurability (W9-A).** `tblLook` first-row flag now
  routes into `<thead>` even without an explicit `<w:tblHeader/>`; web
  server port is configurable via `PORT` env for CI parity.

### Merged branches

- `fix/w1-f-conformance-gaps-2026-05-04`
- `fix/w9-a-conformance-fixes`
- `fix/w9-d-a11y-docxjs`
- `fix/w9-e-responsive`
- `fix/w9-f-shared-css`

### Known

- 11 pre-existing render-test unit failures remain (tracked for
  follow-up; they do not affect the conformance or interop suites).
- Conformance: 263 passing / 43 failing (306 fixtures total).
