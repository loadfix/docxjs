# Interop visual-diff harness

Playwright-driven screenshot regression suite for docxjs, running against
the feature manifests in the sibling
[`ooxml-reference-corpus`](https://github.com/loadfix/ooxml-reference-corpus)
checkout.

## What it does

For every manifest under `../ooxml-reference-corpus/features/docx/*.json`,
the runner:

1. Resolves the manifest's `fixtures.machine` to a `.docx` path under
   `../ooxml-reference-corpus/fixtures/`.
2. Opens the docxjs demo page (`http://localhost:8765/`) and uploads the
   fixture through the `#files` input, which triggers `renderDocx`.
3. Waits for `.docx-wrapper` to appear, then screenshots it.
4. Compares against the committed baseline PNG at
   `tests/interop/baselines/<feature-name>.png`.

The runner is **separate** from the existing `npm test` suite
(`tests/render-test/`) — it has its own Playwright config
(`playwright.interop.config.ts`) and its own report directory
(`playwright-report-interop/`).

## Prerequisites

1. The sibling corpus must exist on disk:

   ```
   git clone https://github.com/loadfix/ooxml-reference-corpus \
       ../ooxml-reference-corpus
   ```

   If it's missing, the harness skips every test with a pointer to this
   instruction.

2. A working Chrome install — the config uses `channel: 'chrome'`, same as
   the main Playwright config.

3. A fresh build of `dist/docx-preview.js`. The demo page loads the UMD
   bundle directly; stale output will produce stale screenshots:

   ```
   npm run build
   ```

## Running

Normal compare-to-baseline run:

```
npm run test:interop
```

Update baselines after a deliberate rendering change:

```
npm run test:interop -- --update-snapshots
```

Filter to a single manifest:

```
npm run test:interop -- -g bold-text
```

Open the HTML report from the last run:

```
npx playwright show-report playwright-report-interop
```

## First-time workflow

On a fresh checkout the `tests/interop/baselines/` directory ships empty
(only a `.gitkeep`). The first `npm run test:interop` will fail every test
— that's expected. You have two options:

- **Review-first (recommended):** run `--update-snapshots`, then
  **visually inspect each PNG** in `tests/interop/baselines/` before
  committing. Every baseline is an assertion about what "correct
  rendering" looks like for that feature, and the tool cannot tell you
  whether the current output is actually correct — only whether it
  matches a previous snapshot.
- **Trust-current:** run `--update-snapshots` and commit without review.
  Use only when you already have high confidence in the current
  rendering (e.g. you just manually QA'd every fixture).

After baselines are committed, subsequent runs will flag any pixel diffs
above the configured threshold.

## Interpreting failures

When a test fails, Playwright writes three images under
`test-results-interop/<test-name>/`:

- `<name>-actual.png` — what the renderer produced this run.
- `<name>-expected.png` — the committed baseline.
- `<name>-diff.png` — a difference map highlighting the pixels that
  changed.

Review the diff. Three common causes:

- **Genuine regression.** Fix the renderer; leave the baseline alone.
- **Intentional rendering change.** Update the baseline with
  `--update-snapshots` and commit it alongside the code change so the
  reason is obvious from the diff.
- **Environmental drift.** See "CI considerations" below.

## Tolerance

Configured in `playwright.interop.config.ts`:

- `threshold: 0.001` — per-pixel colour distance (0..1).
- `maxDiffPixels: 100` — absolute cap on differing pixels.

These are tuned to absorb anti-aliasing noise while still surfacing real
layout or colour regressions. If you find yourself bumping them to paper
over environmental diffs, prefer recommitting baselines from the target
environment instead.

## CI considerations

Screenshot tests are **not portable across environments**. Specifically:

- **OS-dependent text rasterisation.** Fonts rendered on Linux, macOS,
  and Windows differ enough that a Linux-generated baseline will fail on
  macOS and vice versa. Commit baselines generated on the same OS your
  CI runs on — typically Ubuntu.
- **Font availability.** docxjs consumes whatever fonts the browser has
  access to. This harness does not ship fonts, so text-heavy fixtures
  may render differently between a developer's laptop (which has Arial,
  Calibri, Times New Roman) and a minimal CI runner (which may
  substitute DejaVu). Options:
  - Install the expected fonts in CI (`fonts-liberation`,
    `fonts-dejavu`, `fontconfig`).
  - Accept that some fixtures will need CI-local baselines that don't
    match what developers see locally.
- **Chrome version.** Pin `channel: 'chrome'` *and* pin the Chrome
  version in CI; a Chrome major-version upgrade can shift sub-pixel
  anti-aliasing.

Until the font and OS story is nailed down, treat this harness as a
development aid first and a blocking CI gate second. Baselines committed
before that happens will churn as developers run on different hosts.

## Adding a feature

You don't add tests here directly — they're discovered from the corpus.
To add coverage:

1. Contribute a manifest + fixture to `ooxml-reference-corpus`.
2. Pull the corpus update into your sibling checkout.
3. Run `npm run test:interop -- --update-snapshots -g <feature-name>` to
   generate the baseline.
4. Visually inspect the baseline and commit it.
