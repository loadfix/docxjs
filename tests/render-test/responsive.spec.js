// @ts-check
import { test, expect } from '@playwright/test';

// Exercises the `responsive: true` render option added in Wave 9 (W9-E).
//
// Default mode (off) keeps the historical pixel-perfect page width on every
// <section>, so `width` in the inline style must still be set. Responsive
// mode drops that inline width in favour of a fluid `max-width: 100%` rule
// emitted by renderDefaultStyle(), plus a @media (max-width: 768px) block
// that demotes wp:anchor floats to block layout.
//
// Uses the `text/document.docx` fixture because it's the smallest one that
// carries flowing paragraph text.

test.describe('responsive mode', () => {
    test('default (responsive: false): page section keeps inline pixel width', async ({ page }) => {
        await page.goto('/tests/harness.html');

        const result = await page.evaluate(async () => {
            const blob = await fetch('/tests/render-test/text/document.docx').then(r => r.blob());
            const div = document.createElement('div');
            document.body.appendChild(div);
            // @ts-ignore — UMD global
            await docx.renderAsync(blob, div);
            const section = /** @type {HTMLElement} */ (div.querySelector('section.docx'));
            const styleNode = /** @type {HTMLStyleElement} */ (div.querySelector('style'));
            const out = {
                sectionInlineWidth: section?.style.width ?? '',
                cssHasFluidRule: styleNode?.textContent?.includes('max-width: 100%') ?? false,
                cssHasMediaQuery: styleNode?.textContent?.includes('@media (max-width: 768px)') ?? false,
            };
            div.remove();
            return out;
        });

        // The historical inline width is a pt string like "612pt" — just
        // assert it's non-empty rather than binding to a specific page size.
        expect(result.sectionInlineWidth).not.toBe('');
        // Default mode: the new fluid + mobile CSS blocks must NOT be
        // present, so we don't regress byte-stable snapshots for callers
        // who never opt in.
        expect(result.cssHasFluidRule).toBe(false);
        expect(result.cssHasMediaQuery).toBe(false);
    });

    test('responsive: true — page section drops inline width, emits fluid + @media CSS', async ({ page }) => {
        await page.goto('/tests/harness.html');

        const result = await page.evaluate(async () => {
            const blob = await fetch('/tests/render-test/text/document.docx').then(r => r.blob());
            const div = document.createElement('div');
            document.body.appendChild(div);
            // @ts-ignore — UMD global
            await docx.renderAsync(blob, div, undefined, { responsive: true });
            const section = /** @type {HTMLElement} */ (div.querySelector('section.docx'));
            // Collect the textContent across every <style> node the renderer
            // emits (default style, theme, user styles). The responsive
            // block is appended to the default stylesheet but easier to
            // match without being brittle to emission order.
            const cssText = [...div.querySelectorAll('style')]
                .map((s) => s.textContent ?? '')
                .join('\n');
            const out = {
                sectionInlineWidth: section?.style.width ?? '',
                // min-height (page height) must still be set — responsive
                // mode only drops the width, not the height.
                sectionInlineMinHeight: section?.style.minHeight ?? '',
                cssHasFluidRule: cssText.includes('max-width: 100%'),
                cssHasMediaQuery: cssText.includes('@media (max-width: 768px)'),
                cssHasImageRule: cssText.includes('.docx img'),
                cssHasAnchorRule: cssText.includes('[data-drawing-anchor="true"]'),
            };
            div.remove();
            return out;
        });

        expect(result.sectionInlineWidth).toBe('');
        expect(result.sectionInlineMinHeight).not.toBe('');
        expect(result.cssHasFluidRule).toBe(true);
        expect(result.cssHasMediaQuery).toBe(true);
        expect(result.cssHasImageRule).toBe(true);
        expect(result.cssHasAnchorRule).toBe(true);
    });
});
