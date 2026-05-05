// @ts-check
//
// W9-F shared-class assertions — docxjs side.
//
// The renderer now emits a cross-format `oox-*` class alongside each
// format-specific class so manifests can write a single selector that
// works for docxjs/pptxjs/xlsxjs output. This spec verifies the classes
// are present on the expected elements for a representative fixture.
// The per-fixture string-equality specs in test.spec.js strip the
// `oox-*` classes before comparison, so they coexist without edits to
// the golden HTML.
//
// The corresponding renderer logic lives in src/shared-classes.ts
// (table of concept → shared class) and the calls in html-renderer.ts
// next to each `renderParagraph` / `renderRun` / `renderTable` etc.

import { test, expect } from '@playwright/test';

test.describe('Shared oox-* classes', () => {
    test('text fixture carries oox-wrapper / oox-page / oox-paragraph / oox-run', async ({ page }) => {
        await page.goto('/tests/harness.html');

        const counts = await page.evaluate(async () => {
            const docBlob = await fetch('/tests/render-test/text/document.docx').then(r => r.blob());
            const div = document.createElement('div');
            document.body.appendChild(div);
            // @ts-ignore — UMD global
            await docx.renderAsync(docBlob, div);
            const q = (sel) => div.querySelectorAll(sel).length;
            const result = {
                wrapper:    q('.oox-wrapper'),
                page:       q('.oox-page'),
                paragraph:  q('.oox-paragraph'),
                run:        q('.oox-run'),
                // Legacy format-specific class still present (back-compat).
                docxWrapper: q('.docx-wrapper'),
            };
            div.remove();
            return result;
        });

        // The `text` fixture has a wrapper, at least one page, many
        // paragraphs, and many runs. Exact counts depend on the fixture,
        // so we only check lower bounds — what matters is that the
        // classes are emitted at all.
        expect(counts.wrapper).toBeGreaterThanOrEqual(1);
        expect(counts.page).toBeGreaterThanOrEqual(1);
        expect(counts.paragraph).toBeGreaterThanOrEqual(1);
        expect(counts.run).toBeGreaterThanOrEqual(1);
        // Legacy class still emitted.
        expect(counts.docxWrapper).toBeGreaterThanOrEqual(1);
    });

    test('table fixture carries oox-table / oox-table-row / oox-table-cell', async ({ page }) => {
        await page.goto('/tests/harness.html');

        const counts = await page.evaluate(async () => {
            const docBlob = await fetch('/tests/render-test/table/document.docx').then(r => r.blob());
            const div = document.createElement('div');
            document.body.appendChild(div);
            // @ts-ignore — UMD global
            await docx.renderAsync(docBlob, div);
            const q = (sel) => div.querySelectorAll(sel).length;
            const result = {
                table:     q('.oox-table'),
                tableRow:  q('.oox-table-row'),
                tableCell: q('.oox-table-cell'),
            };
            div.remove();
            return result;
        });

        expect(counts.table).toBeGreaterThanOrEqual(1);
        expect(counts.tableRow).toBeGreaterThanOrEqual(1);
        expect(counts.tableCell).toBeGreaterThanOrEqual(1);
    });
});
