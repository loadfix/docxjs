// @ts-check
import { test, expect } from '@playwright/test';

const renderCases = [
    'text',
    'underlines',
    'text-break',
    'table',
    'page-layout',
    'revision',
    'numbering',
    'line-spacing',
    'header-footer',
    'footnote',
    'equation',
];

test.describe('Render document', () => {
    for (const path of renderCases) {
        test(`from ${path} should be correct`, async ({ page }) => {
            await page.goto('/tests/harness.html');

            const { actual, expected } = await page.evaluate(async (p) => {
                const docBlob = await fetch(`/tests/render-test/${p}/document.docx`).then(r => r.blob());
                const resultText = await fetch(`/tests/render-test/${p}/result.html`).then(r => r.text());

                const div = document.createElement('div');
                document.body.appendChild(div);

                // @ts-ignore — `docx` is exposed as a UMD global by dist/docx-preview.js
                await docx.renderAsync(docBlob, div);

                // Strip the cross-format `oox-*` shared classes before
                // comparing to the golden HTML. Those classes are emitted
                // in addition to the format-specific `docx_*` / `docx-*`
                // classes (W9-F shared-CSS consistency work); the goldens
                // predate them and remain the source of truth for the
                // format-specific surface. A dedicated spec
                // (shared-classes.spec.js) asserts the shared classes are
                // present.
                const stripOox = (html) => html
                    // Drop any `oox-*` token (and the preceding space if it
                    // was joining other classes in the same class= value).
                    .replace(/\s*\boox-[\w-]+\b/g, '')
                    // After strip, a lone `class=""` may be left on spans /
                    // <p>s whose only class was the shared one. Remove the
                    // attribute (and the preceding space) so the golden
                    // HTML — which never had a class attribute there —
                    // still matches.
                    .replace(/\s+class=""/g, '');
                const format = (text) => stripOox(text).replace(/\t+|\s+/ig, ' ').replace(/></ig, '>\n<');
                const actual = format(div.innerHTML);
                const expected = format(resultText);

                if (actual !== expected) {
                    // @ts-ignore — `Diff` is exposed as a UMD global by node_modules/diff/dist/diff.js
                    const diffs = Diff.diffLines(expected, actual);
                    for (const d of diffs) {
                        if (d.added) console.log('[+] ' + d.value);
                        if (d.removed) console.log('[-] ' + d.value);
                    }
                }

                div.remove();
                return { actual, expected };
            }, path);

            expect(actual).toBe(expected);
        });
    }
});
