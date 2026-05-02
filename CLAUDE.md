# docxjs — project notes for Claude

Browser-side DOCX→HTML renderer. TypeScript, built with rollup, tested with Karma+jasmine.

## Build & test

- `npm run build` — dev UMD bundle (`dist/docx-preview.js`). The only output rebuilt on every run.
- `npm run build-prod` — also emits `dist/docx-preview.mjs` and minified versions. **Only prod emits the `.mjs`.** If you edit source and load `.mjs` directly for testing, you'll be running stale code.
- `npm run e2e` — Karma suite against real Chrome. Requires a working Chrome launcher in the environment.
- `npm run test:track-changes` — jsdom smoke test for the track-changes pipeline (`scripts/test-track-changes.mjs`). Fast, runs without a browser. Builds the UMD first.

## Testing UI changes

Follow this order for any change that affects rendering:

1. **`npm run build`** — catches type errors.
2. **`npm run test:track-changes` (or equivalent)** — jsdom smoke test. The existing harness covers the track-changes pipeline; for other features, write a similar script that loads the UMD bundle in jsdom, renders a fixture from `tests/render-test/`, and asserts on the DOM. **This is the required minimum** — anything that passes build but fails here is a real regression.
3. **Playwright MCP (when available)** — real-browser confirmation. Required for: visual correctness (colours, positioning, legend layout), event wiring (clicks through the delegated handlers, hover behaviour), layout that depends on real computed styles (sticky positioning, flex measurements). jsdom won't catch these.

### Browser-test recipe

```bash
npm run serve &                  # Node static server on :8765 (PORT=… to override)
#   — or `npm run dev` to rebuild first
# then, via Playwright MCP:
#   browser_navigate → http://localhost:8765/
#   browser_file_upload → tests/render-test/<fixture>/document.docx
#     (uploaded to the `#files` input; fileInput.change fires renderDocx)
#   browser_evaluate → set docxOptions.*, call renderDocx(currentDocument)
#   browser_evaluate → querySelectorAll and assert
#   browser_take_screenshot for visual evidence
#   kill %1 when done
```

The `index.html` demo exposes `docxOptions`, `renderDocx`, and `currentDocument` as globals, so the Playwright flow is driven almost entirely through `browser_evaluate` once a file is loaded.

Fixtures live in `tests/render-test/<name>/document.docx`. The demo has no in-UI fixture picker — load a DOCX via the file input or drag-drop (users of a production viewer shouldn't see test scaffolding).

## Architecture touch points

- `src/document-parser.ts` — DOCX XML → internal element tree. Add new element types here.
- `src/html-renderer.ts` — element tree → DOM nodes. Big file; rendering is a `switch` on `DomType` in `renderElement`.
- `src/document/dom.ts` — the `DomType` enum and `OpenXmlElement` base type. Extend both when adding a new element category.
- `src/docx-preview.ts` — public API surface. `Options`, `defaultOptions`, `parseAsync`, `renderDocument`, `renderAsync`.

### Back-compat patterns used

- `renderChanges: boolean` is the legacy track-changes switch; `changes.show` is the newer nested option. `mergeOptions()` in `docx-preview.ts` translates the legacy flag if `changes.show` isn't set explicitly. **Never drop `renderChanges` from the `Options` interface.**
- Same pattern for `renderComments` vs `comments.*`.

## Security constraints

All string content in a DOCX is attacker-controlled. Treat `author`, `id`, `paraId`, revision ids, comment ids, etc. as untrusted.

- **Never** interpolate DOCX-derived strings into a CSS class name, CSS selector, or `innerHTML`.
- Keyed maps on DOCX-derived strings use either `Map` (safe) or are validated against `SAFE_PARA_ID = /^[A-Za-z0-9_-]+$/` in `src/comments/comments-part.ts`. When adding a new lookup keyed on DOCX data, follow one of those two patterns.
- Attribute values via `setAttribute` or `dataset.*` are safe — the browser HTML-encodes them. That's the preferred sink for DOCX strings.
- Styling on DOCX-derived data goes through a numeric palette index, never the raw string (e.g., `docx-change-author-${getAuthorIndex(author)}`, not `docx-change-author-${author}`).

## Workflow

- Feature branches: `feat/<short-name>`. Test/tooling: `test/<short-name>`.
- Open one PR per logical change. Large features can stack PRs, but GitHub auto-closes stacked PRs when the base branch is deleted on merge — expect to recreate intermediate PRs if you merge bottom-up.
- `.github/workflows/` is intentionally empty (removed 2026-05-01). No automated checks run on PRs; rely on local build + harness + Playwright.
- `dist/` is committed. Rebuild before any commit that touches `src/`.
