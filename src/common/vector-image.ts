// Converts embedded Word vector-image formats (WMF / EMF) into something a
// browser can actually render. Word happily stores legacy `.wmf` / `.emf`
// bitmaps in `/word/media/`, but no browser implements them natively — the
// default `URL.createObjectURL(blob)` path produces broken-image icons.
//
// We route those extensions through this helper. If the host page has loaded
// the `rtf.js` bundle (exposing a `WMFJS` and/or `EMFJS` global, the same way
// consumers already opt into `jszip`), we use its `Renderer` to decode the
// blob into SVG. Otherwise we emit a hard-coded placeholder SVG so the user
// at least sees that a vector image was present.
//
// Security:
// - We never interpolate DOCX-derived strings into the placeholder or the
//   decoder call. Format is validated against the literal strings `"wmf"` /
//   `"emf"` below.
// - Decoded SVG from WMFJS / EMFJS is third-party output we don't fully
//   trust. It's wrapped in a `Blob` of type `image/svg+xml` and exposed as
//   an object URL — never inserted into the DOM as `innerHTML`. The browser
//   sandboxes SVGs loaded via object URL in an `<img>` / CSS `url(...)`
//   context, which neutralises any residual script content in the output.

type VectorFormat = 'wmf' | 'emf';

// Default viewport for the placeholder — matches typical thumbnail sizing.
const PLACEHOLDER_WIDTH = 200;
const PLACEHOLDER_HEIGHT = 100;

// Nominal canvas for decoded output. WMF / EMF records are device-independent,
// so the renderer needs a target extent; we pick a large-ish default and let
// the downstream `<img>` scale to fit its container.
const DECODE_WIDTH = 1000;
const DECODE_HEIGHT = 800;

export async function convertVectorImage(blob: Blob, format: VectorFormat): Promise<string> {
    const decoder = resolveDecoder(format);

    if (decoder) {
        try {
            const svg = await decodeToSvg(blob, decoder);
            if (svg) return toObjectURL(svg);
        } catch {
            // Fall through to the placeholder on any decoder error — a broken
            // or unsupported record shouldn't crash rendering.
        }
    }

    return toObjectURL(placeholderSvg());
}

interface VectorDecoder {
    Renderer: new (buffer: ArrayBuffer | Uint8Array) => {
        render(opts: Record<string, unknown>): Element | { firstChild: Element };
    };
}

function resolveDecoder(format: VectorFormat): VectorDecoder | null {
    const g = globalThis as any;
    if (format === 'wmf' && g.WMFJS?.Renderer) return g.WMFJS;
    if (format === 'emf' && g.EMFJS?.Renderer) return g.EMFJS;
    return null;
}

async function decodeToSvg(blob: Blob, decoder: VectorDecoder): Promise<string | null> {
    const buffer = await blob.arrayBuffer();
    // Some versions of rtf.js accept a Uint8Array, others want the raw
    // ArrayBuffer. Passing a Uint8Array works in both cases.
    const renderer = new decoder.Renderer(new Uint8Array(buffer));
    const result = renderer.render({
        width: `${DECODE_WIDTH}px`,
        height: `${DECODE_HEIGHT}px`,
        xExt: DECODE_WIDTH,
        yExt: DECODE_HEIGHT,
        mapMode: 8, // preserve aspect ratio
    });

    // Older rtf.js returns a wrapper whose firstChild is the <svg>; newer
    // builds return the <svg> directly. Handle both.
    const svgEl: Element | null =
        (result as any)?.tagName?.toLowerCase() === 'svg'
            ? (result as Element)
            : ((result as any)?.firstChild ?? null);

    if (!svgEl || (svgEl as any).tagName?.toLowerCase() !== 'svg') {
        return null;
    }

    // Normalise so the decoded SVG scales nicely inside whatever <img>
    // container the renderer places it in.
    svgEl.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
    svgEl.removeAttribute('width');
    svgEl.removeAttribute('height');

    return (svgEl as any).outerHTML ?? null;
}

function placeholderSvg(): string {
    // Hard-coded — no DOCX data flows into this template.
    return [
        '<svg xmlns="http://www.w3.org/2000/svg"',
        ` viewBox="0 0 ${PLACEHOLDER_WIDTH} ${PLACEHOLDER_HEIGHT}"`,
        ` width="${PLACEHOLDER_WIDTH}" height="${PLACEHOLDER_HEIGHT}">`,
        `<rect x="0" y="0" width="${PLACEHOLDER_WIDTH}" height="${PLACEHOLDER_HEIGHT}"`,
        ' fill="#f0f0f0" stroke="#ccc"/>',
        `<text x="${PLACEHOLDER_WIDTH / 2}" y="${PLACEHOLDER_HEIGHT / 2}"`,
        ' text-anchor="middle" dominant-baseline="middle"',
        ' font-family="sans-serif" font-size="12" fill="#666">',
        'WMF/EMF image',
        '</text>',
        `<text x="${PLACEHOLDER_WIDTH / 2}" y="${PLACEHOLDER_HEIGHT / 2 + 20}"`,
        ' text-anchor="middle" dominant-baseline="middle"',
        ' font-family="sans-serif" font-size="9" fill="#999">',
        '(decoder not loaded)',
        '</text>',
        '</svg>',
    ].join('');
}

function toObjectURL(svg: string): string {
    const prelude = '<?xml version="1.0" encoding="UTF-8"?>';
    const blob = new Blob([prelude, svg], { type: 'image/svg+xml' });
    return URL.createObjectURL(blob);
}

export function detectVectorFormat(path: string | null | undefined): VectorFormat | null {
    if (!path) return null;
    const m = path.toLowerCase().match(/\.([a-z0-9]+)$/);
    const ext = m?.[1];
    return ext === 'wmf' || ext === 'emf' ? ext : null;
}
