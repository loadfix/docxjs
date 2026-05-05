// Render-assertion evaluator.
//
// A thin, side-effect-free layer that knows how to translate a single
// `render_assertion` from an ooxml-reference-corpus manifest into a
// {status, detail} verdict. Three kinds are supported:
//
//   css_selector  — a selector hits some DOM nodes (exist / absent /
//                   equal-count / match-text).
//   computed_style — the first hit's computed style regex-matches value.
//   visual_ssim    — a Playwright screenshot of `.docx-wrapper` is
//                    pixel-diffed against a reference PNG from the
//                    corpus. The "SSIM" score here is a pragmatic
//                    pixelmatch-based (1 - diffRatio); true SSIM is
//                    heavier and unnecessary for go/no-go signal.
//
// Evaluators that inspect the DOM run inside the Playwright page
// (`page.evaluate`) so they see the post-render reality. The SSIM
// evaluator stays on the Node side because that's where the reference
// PNG is on disk and where pixelmatch runs.
//
// Matches the `AssertionResult` shape in
// ooxml-validate/src/ooxml_validate/conformance.py:
//   { id, status: "pass" | "fail" | "error", detail: string }

import { existsSync, readFileSync } from 'node:fs';
import type { Page } from '@playwright/test';

export type Status = 'pass' | 'fail' | 'error';

export interface AssertionVerdict {
    id: string;
    status: Status;
    detail: string;
}

export interface CssSelectorAssertion {
    id: string;
    kind: 'css_selector';
    selector: string;
    must: 'exist' | 'absent' | 'equal-count' | 'match-text';
    count?: number;
    value?: string;
    description?: string;
}

export interface ComputedStyleAssertion {
    id: string;
    kind: 'computed_style';
    selector: string;
    style_property: string;
    value: string;
    description?: string;
}

export interface VisualSsimAssertion {
    id: string;
    kind: 'visual_ssim';
    min_ssim: number;
    reference_png?: string;
    description?: string;
}

export type RenderAssertion =
    | CssSelectorAssertion
    | ComputedStyleAssertion
    | VisualSsimAssertion;

/**
 * Evaluate a css_selector assertion inside the page. All four `must`
 * modes resolve to a single round-trip; selectors that throw (invalid
 * CSS, :has fallback on older browsers) come back as `error`.
 */
export async function evaluateCssSelector(
    page: Page,
    assertion: CssSelectorAssertion,
): Promise<AssertionVerdict> {
    try {
        const outcome = await page.evaluate(
            ({ selector, must, count, value }) => {
                let nodes: NodeListOf<Element>;
                try {
                    nodes = document.querySelectorAll(selector);
                } catch (e) {
                    return { ok: false, error: `querySelectorAll failed: ${(e as Error).message}` };
                }
                const n = nodes.length;
                if (must === 'exist') {
                    return {
                        ok: true,
                        pass: n >= 1,
                        detail: `${n} node(s) matched.`,
                    };
                }
                if (must === 'absent') {
                    return {
                        ok: true,
                        pass: n === 0,
                        detail: n === 0
                            ? 'No nodes matched (as required).'
                            : `${n} node(s) matched but none expected.`,
                    };
                }
                if (must === 'equal-count') {
                    const expected = typeof count === 'number' ? count : -1;
                    return {
                        ok: true,
                        pass: n === expected,
                        detail: `expected ${expected}, got ${n}.`,
                    };
                }
                if (must === 'match-text') {
                    if (n === 0) {
                        return { ok: true, pass: false, detail: 'No node to match against.' };
                    }
                    const text = nodes[0].textContent ?? '';
                    let re: RegExp;
                    try {
                        re = new RegExp(value ?? '');
                    } catch (e) {
                        return { ok: false, error: `Invalid regex ${JSON.stringify(value)}: ${(e as Error).message}` };
                    }
                    const match = re.test(text);
                    return {
                        ok: true,
                        pass: match,
                        detail: match
                            ? `text=${JSON.stringify(text.slice(0, 120))} matches ${JSON.stringify(value)}.`
                            : `text=${JSON.stringify(text.slice(0, 120))} does not match ${JSON.stringify(value)}.`,
                    };
                }
                return { ok: false, error: `Unknown must: ${must}` };
            },
            {
                selector: assertion.selector,
                must: assertion.must,
                count: assertion.count,
                value: assertion.value,
            },
        );
        if (!outcome.ok) {
            return { id: assertion.id, status: 'error', detail: outcome.error };
        }
        return {
            id: assertion.id,
            status: outcome.pass ? 'pass' : 'fail',
            detail: outcome.detail,
        };
    } catch (e) {
        return {
            id: assertion.id,
            status: 'error',
            detail: `Evaluator threw: ${(e as Error).message}`,
        };
    }
}

/**
 * Evaluate a computed_style assertion: find the first selector hit,
 * read getComputedStyle(el).getPropertyValue(style_property), regex-
 * match against value. Empty selection is a fail (nothing to check).
 */
export async function evaluateComputedStyle(
    page: Page,
    assertion: ComputedStyleAssertion,
): Promise<AssertionVerdict> {
    try {
        const outcome = await page.evaluate(
            ({ selector, property, value }) => {
                let el: Element | null;
                try {
                    el = document.querySelector(selector);
                } catch (e) {
                    return { ok: false, error: `querySelector failed: ${(e as Error).message}` };
                }
                if (!el) {
                    return { ok: true, pass: false, detail: `selector ${selector} matched no nodes.` };
                }
                const actual = window.getComputedStyle(el).getPropertyValue(property);
                let re: RegExp;
                try {
                    re = new RegExp(value);
                } catch (e) {
                    return { ok: false, error: `Invalid regex ${JSON.stringify(value)}: ${(e as Error).message}` };
                }
                const match = re.test(actual);
                return {
                    ok: true,
                    pass: match,
                    detail: match
                        ? `${property}=${JSON.stringify(actual)} matches ${JSON.stringify(value)}.`
                        : `${property}=${JSON.stringify(actual)} does not match ${JSON.stringify(value)}.`,
                };
            },
            {
                selector: assertion.selector,
                property: assertion.style_property,
                value: assertion.value,
            },
        );
        if (!outcome.ok) {
            return { id: assertion.id, status: 'error', detail: outcome.error };
        }
        return {
            id: assertion.id,
            status: outcome.pass ? 'pass' : 'fail',
            detail: outcome.detail,
        };
    } catch (e) {
        return {
            id: assertion.id,
            status: 'error',
            detail: `Evaluator threw: ${(e as Error).message}`,
        };
    }
}

/**
 * Evaluate a visual_ssim assertion. Screenshots `.docx-wrapper`, loads
 * the reference PNG from disk, and computes 1 - (diffPixels/total) via
 * pixelmatch as a pragmatic similarity score. Size mismatches between
 * the docxjs screenshot and the PDF-derived reference return `error`
 * with a detail rather than failing the comparison, because they carry
 * no real signal.
 */
export async function evaluateVisualSsim(
    page: Page,
    assertion: VisualSsimAssertion,
    referencePngAbsPath: string,
): Promise<AssertionVerdict> {
    if (!existsSync(referencePngAbsPath)) {
        return {
            id: assertion.id,
            status: 'error',
            detail: `reference PNG not found: ${referencePngAbsPath}`,
        };
    }
    // Capture the wrapper. Locator-scoped screenshot crops to the
    // bounding box, so the reference PNG needs roughly the same crop.
    let screenshotBuffer: Buffer;
    try {
        const wrapper = page.locator('.docx-wrapper').first();
        await wrapper.waitFor({ timeout: 5_000 });
        screenshotBuffer = await wrapper.screenshot();
    } catch (e) {
        return {
            id: assertion.id,
            status: 'error',
            detail: `Could not screenshot .docx-wrapper: ${(e as Error).message}`,
        };
    }

    // pixelmatch is pure-ESM (type: module), pngjs is CJS. Use dynamic
    // import for pixelmatch so it works under Playwright's CJS test
    // loader; require() it would throw ERR_REQUIRE_ESM.
    let pixelmatch: typeof import('pixelmatch').default;
    let PNG: typeof import('pngjs').PNG;
    try {
        pixelmatch = (await import('pixelmatch')).default;
        // pngjs exports PNG as a property on the module object under CJS.
        const pngMod = await import('pngjs');
        PNG = (pngMod as any).PNG ?? (pngMod as any).default?.PNG;
    } catch (e) {
        return {
            id: assertion.id,
            status: 'error',
            detail: `Could not load pixelmatch/pngjs: ${(e as Error).message}`,
        };
    }

    let refPng: InstanceType<typeof PNG>;
    let gotPng: InstanceType<typeof PNG>;
    try {
        refPng = PNG.sync.read(readFileSync(referencePngAbsPath));
        gotPng = PNG.sync.read(screenshotBuffer);
    } catch (e) {
        return {
            id: assertion.id,
            status: 'error',
            detail: `Could not decode PNG: ${(e as Error).message}`,
        };
    }

    // If dimensions don't match, bilinearly resize the screenshot to
    // the reference dimensions. The reference PNG is usually the
    // PDF-derived LibreOffice render (letter @ 150dpi = 1275x1650);
    // the docxjs screenshot is the .docx-wrapper crop at whatever the
    // Playwright viewport + deviceScaleFactor produces. A resize loses
    // some fidelity but keeps SSIM as a usable signal, whereas erroring
    // out on size mismatch turns every visual assertion into noise.
    if (refPng.width !== gotPng.width || refPng.height !== gotPng.height) {
        try {
            const resized = resizePngBilinear(gotPng, refPng.width, refPng.height, PNG);
            // Overwrite gotPng's data/width/height with the resized
            // version so the pixelmatch call below sees matching
            // dimensions.
            (gotPng as any).data = resized.data;
            (gotPng as any).width = resized.width;
            (gotPng as any).height = resized.height;
        } catch (e) {
            return {
                id: assertion.id,
                status: 'error',
                detail: `Resize failed (ref ${refPng.width}x${refPng.height}, got ${gotPng.width}x${gotPng.height}): ${(e as Error).message}`,
            };
        }
    }

    const { width, height } = refPng;
    const total = width * height;
    let diff: number;
    try {
        // threshold=0.3 is generous enough to tolerate anti-aliasing
        // and hinting drift between font renderers (Liberation Serif
        // vs Times New Roman baked into the LibreOffice reference PDF)
        // while still catching layout regressions (wrong-width
        // paragraphs, missing borders, misaligned tables). The
        // pixelmatch threshold is a YIQ-space colour distance — 0.1
        // fails on sub-pixel font antialiasing even between two
        // renderings of the same glyph.
        diff = pixelmatch(refPng.data, gotPng.data, undefined, width, height, { threshold: 0.3 });
    } catch (e) {
        return {
            id: assertion.id,
            status: 'error',
            detail: `pixelmatch threw: ${(e as Error).message}`,
        };
    }
    const similarity = 1 - diff / total;
    if (similarity >= assertion.min_ssim) {
        return {
            id: assertion.id,
            status: 'pass',
            detail: `similarity=${similarity.toFixed(4)} >= ${assertion.min_ssim}.`,
        };
    }
    return {
        id: assertion.id,
        status: 'fail',
        detail: `similarity=${similarity.toFixed(4)} < ${assertion.min_ssim}.`,
    };
}

/**
 * Bilinear resize of an RGBA PNG to the given dimensions. Returns a
 * new PNG with a `data` buffer, `width`, `height`. No libraries — pure
 * JS so the conformance runner stays dependency-light.
 *
 * Bilinear is good enough for SSIM: we're comparing visual similarity
 * at the ~85% threshold, not pixel-perfect exactness, and the
 * alternative (mismatched sizes erroring out entirely) is worse.
 */
function resizePngBilinear(
    src: { data: Buffer | Uint8Array; width: number; height: number },
    dstW: number,
    dstH: number,
    PNG: any,
): { data: Buffer; width: number; height: number } {
    const dst = new PNG({ width: dstW, height: dstH });
    const sx = src.width / dstW;
    const sy = src.height / dstH;
    const srcData = src.data as Uint8Array | Buffer;
    for (let y = 0; y < dstH; y++) {
        const fy = (y + 0.5) * sy - 0.5;
        const y0 = Math.max(0, Math.floor(fy));
        const y1 = Math.min(src.height - 1, y0 + 1);
        const wy = fy - y0;
        for (let x = 0; x < dstW; x++) {
            const fx = (x + 0.5) * sx - 0.5;
            const x0 = Math.max(0, Math.floor(fx));
            const x1 = Math.min(src.width - 1, x0 + 1);
            const wx = fx - x0;
            const i00 = (y0 * src.width + x0) << 2;
            const i01 = (y0 * src.width + x1) << 2;
            const i10 = (y1 * src.width + x0) << 2;
            const i11 = (y1 * src.width + x1) << 2;
            const dstIdx = (y * dstW + x) << 2;
            for (let c = 0; c < 4; c++) {
                const v = (1 - wx) * (1 - wy) * srcData[i00 + c]
                        + wx * (1 - wy) * srcData[i01 + c]
                        + (1 - wx) * wy * srcData[i10 + c]
                        + wx * wy * srcData[i11 + c];
                dst.data[dstIdx + c] = Math.round(v);
            }
        }
    }
    return { data: dst.data, width: dstW, height: dstH };
}

/**
 * Dispatch an assertion to the right evaluator. Unknown kinds come
 * back as `error` rather than throwing so the rest of the manifest can
 * still be evaluated.
 */
export async function evaluateAssertion(
    page: Page,
    assertion: RenderAssertion,
    referencePngResolver: (a: VisualSsimAssertion) => string,
): Promise<AssertionVerdict> {
    if (assertion.kind === 'css_selector') {
        return evaluateCssSelector(page, assertion);
    }
    if (assertion.kind === 'computed_style') {
        return evaluateComputedStyle(page, assertion);
    }
    if (assertion.kind === 'visual_ssim') {
        return evaluateVisualSsim(page, assertion, referencePngResolver(assertion));
    }
    const unknown = assertion as { id?: string; kind?: string };
    return {
        id: unknown.id ?? '<unknown>',
        status: 'error',
        detail: `Unknown render_assertion kind: ${unknown.kind}`,
    };
}
