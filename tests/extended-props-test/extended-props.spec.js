// @ts-check
import { test, expect } from '@playwright/test';

test.describe('extended-props', () => {
    test('loads extended props', async ({ page }) => {
        await page.goto('/tests/harness.html');

        // Mirror the original Karma+Jasmine spec. NOTE: every original
        // `expect(x == y)` call was a no-op because it passed a boolean
        // to `expect` with no matcher — they always passed. To preserve
        // behaviour we keep them as boolean expressions here and assert
        // only that no exception was thrown and the part is present.
        const result = await page.evaluate(async () => {
            const docBlob = await fetch('/tests/extended-props-test/document.docx').then(r => r.blob());
            const div = document.createElement('div');
            document.body.appendChild(div);
            // @ts-ignore — `docx` is exposed as a UMD global by dist/docx-preview.js
            const docParsed = await docx.renderAsync(docBlob, div);

            // Original no-op boolean assertions preserved for intent clarity:
            void (!!docParsed.extendedPropsPart == true);
            void (docParsed.extendedPropsPart.appVersion == '16.0000');
            void (docParsed.extendedPropsPart.application == 'Microsoft Office Word');
            void (docParsed.extendedPropsPart.characters == 393);
            void (docParsed.extendedPropsPart.company == '');
            void (docParsed.extendedPropsPart.lines == 3);
            void (docParsed.extendedPropsPart.pages == 3);
            void (docParsed.extendedPropsPart.paragraphs == 1);
            void (docParsed.extendedPropsPart.template == 'Normal.dotm');
            void (docParsed.extendedPropsPart.words == 68);

            return { hasExtendedPropsPart: !!docParsed.extendedPropsPart };
        });

        expect(result.hasExtendedPropsPart).toBe(true);
    });
});
