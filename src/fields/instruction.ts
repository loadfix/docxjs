// Minimal parser for Word field instructions (the string that appears inside
// <w:fldSimple w:instr="…"> or that is assembled from the <w:instrText> runs
// between <w:fldChar fldCharType="begin"/> and <w:fldChar fldCharType="separate"/>).
//
// We only need to recognise the field *code* (PAGE, HYPERLINK, REF, …), the
// backslash switches (\l, \h, …) and the positional arguments. We do NOT
// evaluate the field — the renderer uses the cached result that Word already
// stored inside the field. See html-renderer.ts renderSimpleField /
// renderComplexFieldGroup.

export interface ParsedFieldInstruction {
    code: string;
    switches: string[];
    args: string[];
    raw: string;
}

export function parseFieldInstruction(raw: string): ParsedFieldInstruction {
    const result: ParsedFieldInstruction = {
        code: '',
        switches: [],
        args: [],
        raw: raw ?? '',
    };
    if (!raw) return result;

    const tokens: string[] = [];
    let i = 0;
    const s = raw;
    while (i < s.length) {
        const ch = s[i];
        if (ch === ' ' || ch === '\t' || ch === '\r' || ch === '\n') {
            i++;
            continue;
        }
        if (ch === '"') {
            // Quoted argument — ends at the next unescaped double quote.
            // Backslash escapes the quote character itself.
            let buf = '';
            i++;
            while (i < s.length && s[i] !== '"') {
                if (s[i] === '\\' && i + 1 < s.length && s[i + 1] === '"') {
                    buf += '"';
                    i += 2;
                } else {
                    buf += s[i++];
                }
            }
            if (i < s.length) i++; // consume closing quote
            tokens.push(buf);
            continue;
        }
        // Bare token — runs until the next whitespace. Note: a leading
        // backslash is part of the token so we can surface it as a switch.
        let buf = '';
        while (i < s.length && !/\s/.test(s[i])) {
            buf += s[i++];
        }
        tokens.push(buf);
    }

    if (tokens.length === 0) return result;
    result.code = tokens[0].toUpperCase();
    for (let k = 1; k < tokens.length; k++) {
        const t = tokens[k];
        if (t.startsWith('\\') && t.length > 1) result.switches.push(t);
        else result.args.push(t);
    }
    return result;
}
