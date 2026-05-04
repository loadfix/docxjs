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
#   browser_take_screenshot with filename="./.screenshots/<name>.png"
#   kill %1 when done
```

**Screenshot convention**: always pass `filename="./.screenshots/<descriptive-name>.png"` to `browser_take_screenshot` so images don't pile up in the repo root. `.screenshots/` is gitignored. `.playwright-mcp/` (auto-created by the MCP for console/snapshot logs) is also gitignored — leave it alone.

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

## OOXML feature workflow (required before adding rendering for any new feature)

Every OOXML feature is defined by a manifest in the shared corpus
repository `loadfix/ooxml-reference-corpus` (sibling checkout at
`../ooxml-reference-corpus/`). docxjs is a renderer — it reads a fixture
rather than authoring one — but it must agree with `python-docx` on what
the feature's XML looks like.

1. **Read the manifest.** Look under
   `../ooxml-reference-corpus/features/docx/` for a JSON manifest
   covering the feature you're rendering. The `assertions` block tells
   you what XML to expect in the fixture.

2. **Consult the ECMA-376 5th edition spec** (corpus-only):
   - PDFs: `../ooxml-reference-corpus/spec/ecma-376-5/part-{1,2,3,4}/*.pdf`
   - RNC schemas (easier to read): `../ooxml-reference-corpus/spec/ecma-376-5/part-1/rnc/`
   - XSD schemas: `../ooxml-reference-corpus/spec/ecma-376-5/part-1/xsd/`

3. **If no manifest exists**, ask python-docx's maintainer to author
   one first — the manifest is the definition of "done", and
   authoring-side libraries own the definition. docxjs then confirms
   it can render the committed fixture correctly.

4. **Verify rendering.** The Playwright test suite should include a
   render smoke for the fixture (load the committed fixture from the
   corpus, render, assert the DOM + visual output).

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

### Keep README.md and TODO.md current

Whenever a feature is added, removed, or a public option changes, update both of these files *in the same PR* as the code change — stale docs have bitten us before.

- **`README.md`** — the API block reflects the real `Options` interface. If you add/remove an option (including nested ones like `comments.*` or `changes.*`), add/remove a row in the API code block. If you add or remove an `export` in `src/docx-preview.ts`, reflect it in the API section. The "Breaks" / "Thumbnails" / "Status" prose should also match reality — e.g. if `experimentalPageBreaks` gains a capability, call it out there.
- **`TODO.md`** — if the change resolves an upstream issue listed under "Bugs to investigate" or "New features to consider", move that entry into the "Resolved in fork" table with a one-line description and the PR number. Update the counts table at the top and bump the "last updated" date.

Minimum check before every PR that touches `src/`: `grep -n "<option name>" README.md TODO.md` to catch stale references.

## Parallelising work with subagents

When there are multiple independent tasks (bug fixes on different files, triage of many issues, a set of small low-risk patches), prefer spawning **multiple subagents in parallel** over working through them sequentially. The Agent tool accepts `run_in_background: true` — use it. Typical win: 5 small fixes in ~2 minutes wall-clock instead of 10 minutes sequential.

**When to fan out**:
- The tasks don't depend on each other's outputs.
- Each task is self-contained enough to explain in a short prompt.
- The scope is bounded (you can predict the files touched).

**When NOT to fan out**:
- Tasks are sequential by nature (e.g., "write the parser, then have the renderer use it").
- You need to maintain a single coherent narrative across the work (design docs, major refactors).
- One task's output changes the inputs to another.

### Required: isolated git worktrees

Each subagent MUST work in its own git worktree. Multiple agents operating on the same working tree at `/home/ben/code/docxjs` will silently overwrite each other's edits — observed repeatedly in practice. The pattern each subagent should follow:

```bash
git fetch origin
git worktree add /tmp/docxjs-<task-name> origin/master -b <branch-name>
cd /tmp/docxjs-<task-name>
ln -s /home/ben/code/docxjs/node_modules node_modules   # reuse deps
# ...do work, commit, push...
cd /home/ben/code/docxjs
git worktree remove /tmp/docxjs-<task-name>
```

Include this recipe in every subagent prompt for parallel work.

### Integration pass (yours, not the subagents')

Subagents should **not** rebuild `dist/`, add harness scenarios, or open PRs. Instead, reserve those for a single integration pass after all subagents complete:

1. Pull each branch into an umbrella branch (`chore/follow-ups` or similar).
2. Rebuild `dist/` once.
3. Add any new harness scenarios in one coherent pass (avoids scenario-numbering conflicts — observed earlier in this project's history when multiple branches each tried to claim "scenario 9").
4. Run `npm run test:track-changes` and Playwright verification.
5. Open one combined PR, or merge the subagent PRs in dependency order.

### Prompt template for a fan-out subagent

```
You're fixing <issue> in the docxjs project at /home/ben/code/docxjs.

## Context
Read: <files>
Start by creating an isolated worktree (see CLAUDE.md "Parallelising work
with subagents"). Do all work there. Do NOT touch the main checkout.

## The fix
<scope>

## Constraints
- Do NOT rebuild dist/.
- Do NOT add harness scenarios in scripts/test-track-changes.mjs.
- Do NOT open a PR.

## Verification
npm run build must succeed.

## Commit
One commit on <branch-name>:
<message>

## Report back (under 150 words)
- Files changed.
- Whether npm run build passed.
- Any decisions worth flagging.
```

Keep the "Do NOT" list explicit — agents happily do all three by default, which creates conflicts you then have to untangle.
