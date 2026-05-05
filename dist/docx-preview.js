/*
 * @license
 * docx-preview <https://github.com/VolodymyrBaydalka/docxjs>
 * Released under Apache License 2.0  <https://github.com/VolodymyrBaydalka/docxjs/blob/master/LICENSE>
 * Copyright Volodymyr Baydalka
 */
(function (global, factory) {
    typeof exports === 'object' && typeof module !== 'undefined' ? factory(exports, require('jszip')) :
    typeof define === 'function' && define.amd ? define(['exports', 'jszip'], factory) :
    (global = typeof globalThis !== 'undefined' ? globalThis : global || self, factory(global.docx = {}, global.JSZip));
})(this, (function (exports, JSZip) { 'use strict';

    var RelationshipTypes;
    (function (RelationshipTypes) {
        RelationshipTypes["OfficeDocument"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        RelationshipTypes["FontTable"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
        RelationshipTypes["Image"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        RelationshipTypes["Numbering"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
        RelationshipTypes["Styles"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
        RelationshipTypes["StylesWithEffects"] = "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects";
        RelationshipTypes["Theme"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
        RelationshipTypes["Settings"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
        RelationshipTypes["WebSettings"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings";
        RelationshipTypes["Hyperlink"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
        RelationshipTypes["Footnotes"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
        RelationshipTypes["Endnotes"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
        RelationshipTypes["Footer"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
        RelationshipTypes["Header"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
        RelationshipTypes["ExtendedProperties"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
        RelationshipTypes["CoreProperties"] = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        RelationshipTypes["CustomProperties"] = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/custom-properties";
        RelationshipTypes["Comments"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
        RelationshipTypes["CommentsExtended"] = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
        RelationshipTypes["AltChunk"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk";
        RelationshipTypes["Chart"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
        RelationshipTypes["ChartEx"] = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
        RelationshipTypes["DiagramData"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData";
        RelationshipTypes["DiagramLayout"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout";
        RelationshipTypes["DiagramQuickStyle"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle";
        RelationshipTypes["DiagramColors"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors";
        RelationshipTypes["DiagramDrawing"] = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";
        RelationshipTypes["GlossaryDocument"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument";
        RelationshipTypes["CustomXml"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
        RelationshipTypes["CustomXmlProps"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
    })(RelationshipTypes || (RelationshipTypes = {}));
    function parseRelationships(root, xml) {
        return xml.elements(root).map(e => ({
            id: xml.attr(e, "Id"),
            type: xml.attr(e, "Type"),
            target: xml.attr(e, "Target"),
            targetMode: xml.attr(e, "TargetMode")
        }));
    }

    function escapeClassName(className) {
        return className?.replace(/[ .]+/g, '-').replace(/[&]+/g, 'and').toLowerCase();
    }
    function encloseFontFamily(fontFamily) {
        return /^[^"'].*\s.*[^"']$/.test(fontFamily) ? `'${fontFamily}'` : fontFamily;
    }
    function sanitizeFontFamily(value) {
        if (typeof value !== 'string')
            return 'sans-serif';
        const cleaned = value.replace(/["'\\;{}@<>]/g, '').trim();
        if (!cleaned)
            return 'sans-serif';
        return `'${cleaned}'`;
    }
    const HEX_COLOR_RE = /^[0-9A-Fa-f]{3,8}$/;
    const CSS_FN_COLOR_RE = /^(rgb|rgba|hsl|hsla)\(\s*[-0-9.,%\s/deg]+\s*\)$/i;
    function sanitizeCssColor(value) {
        if (typeof value !== 'string')
            return null;
        const v = value.trim();
        if (!v)
            return null;
        if (HEX_COLOR_RE.test(v))
            return `#${v}`;
        if (v.startsWith('#') && HEX_COLOR_RE.test(v.slice(1)))
            return v;
        if (CSS_FN_COLOR_RE.test(v))
            return v;
        return null;
    }
    const SAFE_CSS_IDENT_RE = /^[A-Za-z0-9_]+$/;
    function isSafeCssIdent(value) {
        return typeof value === 'string' && SAFE_CSS_IDENT_RE.test(value);
    }
    function escapeCssStringContent(value) {
        return value
            .replace(/\\/g, '\\\\')
            .replace(/"/g, '\\"')
            .replace(/\n/g, '\\A ')
            .replace(/\r/g, '\\D ');
    }
    function splitPath(path) {
        let si = path.lastIndexOf('/') + 1;
        let folder = si == 0 ? "" : path.substring(0, si);
        let fileName = si == 0 ? path : path.substring(si);
        return [folder, fileName];
    }
    function resolvePath(path, base) {
        try {
            const prefix = "http://docx/";
            const url = new URL(path, prefix + base).toString();
            return url.substring(prefix.length);
        }
        catch {
            return `${base}${path}`;
        }
    }
    const UNSAFE_KEYS = new Set(['__proto__', 'constructor', 'prototype']);
    function keyBy(array, by) {
        const result = Object.create(null);
        for (const x of array) {
            const k = by(x);
            if (k == null)
                continue;
            const s = String(k);
            if (UNSAFE_KEYS.has(s))
                continue;
            result[s] = x;
        }
        return result;
    }
    function blobToBase64(blob) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve(reader.result);
            reader.onerror = () => reject();
            reader.readAsDataURL(blob);
        });
    }
    function isObject(item) {
        return item && typeof item === 'object' && !Array.isArray(item);
    }
    function isString(item) {
        return typeof item === 'string' || item instanceof String;
    }
    function mergeDeep(target, ...sources) {
        if (!sources.length)
            return target;
        const source = sources.shift();
        if (isObject(target) && isObject(source)) {
            for (const key in source) {
                if (UNSAFE_KEYS.has(key))
                    continue;
                if (!Object.prototype.hasOwnProperty.call(source, key))
                    continue;
                if (isObject(source[key])) {
                    const val = target[key] ?? (target[key] = {});
                    mergeDeep(val, source[key]);
                }
                else {
                    target[key] = source[key];
                }
            }
        }
        return mergeDeep(target, ...sources);
    }
    function parseCssRules(text) {
        const result = {};
        for (const rule of text.split(';')) {
            const [key, val] = rule.split(':');
            result[key] = val;
        }
        return result;
    }
    function asArray(val) {
        return Array.isArray(val) ? val : [val];
    }
    function clamp$1(val, min, max) {
        return min > val ? min : (max < val ? max : val);
    }

    const ns$1 = {
        wordml: "http://schemas.openxmlformats.org/wordprocessingml/2006/main"};
    const LengthUsage = {
        Dxa: { mul: 0.05, unit: "pt" },
        Emu: { mul: 1 / 12700, unit: "pt" },
        FontSize: { mul: 0.5, unit: "pt" },
        Border: { mul: 0.125, unit: "pt", min: 0.25, max: 12 },
        Point: { mul: 1, unit: "pt" },
        Percent: { mul: 0.02, unit: "%" },
        VmlEmu: { mul: 1 / 12700, unit: "" },
    };
    function convertLength(val, usage = LengthUsage.Dxa) {
        if (val == null || /.+(p[xt]|[%])$/.test(val)) {
            return val;
        }
        var num = parseInt(val) * usage.mul;
        if (usage.min && usage.max)
            num = clamp$1(num, usage.min, usage.max);
        return `${num.toFixed(2)}${usage.unit}`;
    }
    function convertBoolean(v, defaultValue = false) {
        switch (v) {
            case "1": return true;
            case "0": return false;
            case "on": return true;
            case "off": return false;
            case "true": return true;
            case "false": return false;
            default: return defaultValue;
        }
    }
    function parseCommonProperty(elem, props, xml) {
        if (elem.namespaceURI != ns$1.wordml)
            return false;
        switch (elem.localName) {
            case "color":
                props.color = xml.attr(elem, "val");
                break;
            case "sz":
                props.fontSize = xml.lengthAttr(elem, "val", LengthUsage.FontSize);
                break;
            default:
                return false;
        }
        return true;
    }

    function parseXmlString(xmlString, trimXmlDeclaration = false) {
        if (trimXmlDeclaration)
            xmlString = xmlString.replace(/<[?].*[?]>/, "");
        xmlString = removeUTF8BOM(xmlString);
        const result = new DOMParser().parseFromString(xmlString, "application/xml");
        const errorText = hasXmlParserError(result);
        if (errorText)
            throw new Error(errorText);
        return result;
    }
    function hasXmlParserError(doc) {
        return doc.getElementsByTagName("parsererror")[0]?.textContent;
    }
    function removeUTF8BOM(data) {
        return data.charCodeAt(0) === 0xFEFF ? data.substring(1) : data;
    }
    function serializeXmlString(elem) {
        return new XMLSerializer().serializeToString(elem);
    }
    class XmlParser {
        elements(elem, localName = null) {
            const result = [];
            for (let i = 0, l = elem.childNodes.length; i < l; i++) {
                let c = elem.childNodes.item(i);
                if (c.nodeType == Node.ELEMENT_NODE && (localName == null || c.localName == localName))
                    result.push(c);
            }
            return result;
        }
        element(elem, localName) {
            for (let i = 0, l = elem.childNodes.length; i < l; i++) {
                let c = elem.childNodes.item(i);
                if (c.nodeType == 1 && c.localName == localName)
                    return c;
            }
            return null;
        }
        elementAttr(elem, localName, attrLocalName) {
            var el = this.element(elem, localName);
            return el ? this.attr(el, attrLocalName) : undefined;
        }
        attrs(elem) {
            return Array.from(elem.attributes);
        }
        attr(elem, localName) {
            for (let i = 0, l = elem.attributes.length; i < l; i++) {
                let a = elem.attributes.item(i);
                if (a.localName == localName)
                    return a.value;
            }
            return null;
        }
        intAttr(node, attrName, defaultValue = null) {
            var val = this.attr(node, attrName);
            return val ? parseInt(val) : defaultValue;
        }
        hexAttr(node, attrName, defaultValue = null) {
            var val = this.attr(node, attrName);
            return val ? parseInt(val, 16) : defaultValue;
        }
        floatAttr(node, attrName, defaultValue = null) {
            var val = this.attr(node, attrName);
            return val ? parseFloat(val) : defaultValue;
        }
        boolAttr(node, attrName, defaultValue = null) {
            return convertBoolean(this.attr(node, attrName), defaultValue);
        }
        lengthAttr(node, attrName, usage = LengthUsage.Dxa) {
            return convertLength(this.attr(node, attrName), usage);
        }
    }
    const globalXmlParser = new XmlParser();

    class Part {
        constructor(_package, path) {
            this._package = _package;
            this.path = path;
        }
        async load() {
            this.rels = await this._package.loadRelationships(this.path);
            const xmlText = await this._package.load(this.path);
            const xmlDoc = this._package.parseXmlDocument(xmlText);
            if (this._package.options.keepOrigin) {
                this._xmlDocument = xmlDoc;
            }
            this.parseXml(xmlDoc.firstElementChild);
        }
        save() {
            this._package.update(this.path, serializeXmlString(this._xmlDocument));
        }
        parseXml(root) {
        }
    }

    const embedFontTypeMap = {
        embedRegular: 'regular',
        embedBold: 'bold',
        embedItalic: 'italic',
        embedBoldItalic: 'boldItalic',
    };
    function parseFonts(root, xml) {
        return xml.elements(root).map(el => parseFont(el, xml));
    }
    function parseFont(elem, xml) {
        let result = {
            name: xml.attr(elem, "name"),
            embedFontRefs: []
        };
        for (let el of xml.elements(elem)) {
            switch (el.localName) {
                case "family":
                    result.family = xml.attr(el, "val");
                    break;
                case "altName":
                    result.altName = xml.attr(el, "val");
                    break;
                case "embedRegular":
                case "embedBold":
                case "embedItalic":
                case "embedBoldItalic":
                    result.embedFontRefs.push(parseEmbedFontRef(el, xml));
                    break;
            }
        }
        return result;
    }
    function parseEmbedFontRef(elem, xml) {
        return {
            id: xml.attr(elem, "id"),
            key: xml.attr(elem, "fontKey"),
            type: embedFontTypeMap[elem.localName]
        };
    }

    class FontTablePart extends Part {
        parseXml(root) {
            this.fonts = parseFonts(root, this._package.xmlParser);
        }
    }

    function parseContentTypes(root, xml) {
        return xml.elements(root).map(e => ({
            extension: xml.attr(e, "Extension"),
            partName: xml.attr(e, "PartName"),
            contentType: xml.attr(e, "ContentType")
        }));
    }

    class OpenXmlPackage {
        constructor(_zip, options) {
            this._zip = _zip;
            this.options = options;
            this.xmlParser = new XmlParser();
        }
        get(path) {
            const p = normalizePath(path);
            return this._zip.files[p] ?? this._zip.files[p.replace(/\//g, '\\')];
        }
        update(path, content) {
            this._zip.file(path, content);
        }
        static async load(input, options) {
            const zip = await JSZip.loadAsync(input);
            return new OpenXmlPackage(zip, options);
        }
        save(type = "blob") {
            return this._zip.generateAsync({ type });
        }
        load(path, type = "string") {
            return this.get(path)?.async(type) ?? Promise.resolve(null);
        }
        async loadRelationships(path = null) {
            let relsPath = `_rels/.rels`;
            if (path != null) {
                const [f, fn] = splitPath(path);
                relsPath = `${f}_rels/${fn}.rels`;
            }
            const txt = await this.load(relsPath);
            return txt ? parseRelationships(this.parseXmlDocument(txt).firstElementChild, this.xmlParser) : null;
        }
        async loadContentTypes() {
            const txt = await this.load("[Content_Types].xml");
            return txt ? parseContentTypes(this.parseXmlDocument(txt).firstElementChild, this.xmlParser) : [];
        }
        parseXmlDocument(txt) {
            return parseXmlString(txt, this.options.trimXmlDeclaration);
        }
    }
    function normalizePath(path) {
        return path.startsWith('/') ? path.substr(1) : path;
    }

    class DocumentPart extends Part {
        constructor(pkg, path, parser) {
            super(pkg, path);
            this._documentParser = parser;
        }
        parseXml(root) {
            this.body = this._documentParser.parseDocumentFile(root);
        }
    }

    function parseBorder(elem, xml) {
        return {
            type: xml.attr(elem, "val"),
            color: xml.attr(elem, "color"),
            size: xml.lengthAttr(elem, "sz", LengthUsage.Border),
            offset: xml.lengthAttr(elem, "space", LengthUsage.Point),
            frame: xml.boolAttr(elem, 'frame'),
            shadow: xml.boolAttr(elem, 'shadow')
        };
    }
    function parseBorders(elem, xml) {
        var result = {};
        for (let e of xml.elements(elem)) {
            switch (e.localName) {
                case "left":
                    result.left = parseBorder(e, xml);
                    break;
                case "top":
                    result.top = parseBorder(e, xml);
                    break;
                case "right":
                    result.right = parseBorder(e, xml);
                    break;
                case "bottom":
                    result.bottom = parseBorder(e, xml);
                    break;
            }
        }
        return result;
    }

    var SectionType;
    (function (SectionType) {
        SectionType["Continuous"] = "continuous";
        SectionType["NextPage"] = "nextPage";
        SectionType["NextColumn"] = "nextColumn";
        SectionType["EvenPage"] = "evenPage";
        SectionType["OddPage"] = "oddPage";
    })(SectionType || (SectionType = {}));
    function parseSectionProperties(elem, xml = globalXmlParser) {
        var section = {};
        for (let e of xml.elements(elem)) {
            switch (e.localName) {
                case "pgSz":
                    section.pageSize = {
                        width: xml.lengthAttr(e, "w"),
                        height: xml.lengthAttr(e, "h"),
                        orientation: xml.attr(e, "orient")
                    };
                    break;
                case "type":
                    section.type = xml.attr(e, "val");
                    break;
                case "pgMar":
                    section.pageMargins = {
                        left: xml.lengthAttr(e, "left"),
                        right: xml.lengthAttr(e, "right"),
                        top: xml.lengthAttr(e, "top"),
                        bottom: xml.lengthAttr(e, "bottom"),
                        header: xml.lengthAttr(e, "header"),
                        footer: xml.lengthAttr(e, "footer"),
                        gutter: xml.lengthAttr(e, "gutter"),
                    };
                    break;
                case "cols":
                    section.columns = parseColumns(e, xml);
                    break;
                case "headerReference":
                    (section.headerRefs ?? (section.headerRefs = [])).push(parseFooterHeaderReference(e, xml));
                    break;
                case "footerReference":
                    (section.footerRefs ?? (section.footerRefs = [])).push(parseFooterHeaderReference(e, xml));
                    break;
                case "titlePg":
                    section.titlePage = xml.boolAttr(e, "val", true);
                    break;
                case "pgBorders":
                    section.pageBorders = parseBorders(e, xml);
                    break;
                case "pgNumType":
                    section.pageNumber = parsePageNumber(e, xml);
                    break;
                case "lnNumType":
                    section.lineNumbering = parseLineNumbering(e, xml);
                    break;
                case "docGrid":
                    section.docGrid = parseDocGrid(e, xml);
                    break;
                case "mirrorMargins":
                    section.mirrorMargins = xml.boolAttr(e, "val", true);
                    break;
            }
        }
        return section;
    }
    function parseLineNumbering(elem, xml) {
        return {
            countBy: xml.intAttr(elem, "countBy", 1),
            start: xml.intAttr(elem, "start", 1),
            distance: xml.lengthAttr(elem, "distance"),
            restart: xml.attr(elem, "restart") || "newPage",
        };
    }
    function parseDocGrid(elem, xml) {
        return {
            type: xml.attr(elem, "type") || "default",
            linePitch: xml.intAttr(elem, "linePitch", 0),
            charSpace: xml.intAttr(elem, "charSpace", 0),
        };
    }
    function parseColumns(elem, xml) {
        return {
            numberOfColumns: xml.intAttr(elem, "num"),
            space: xml.lengthAttr(elem, "space"),
            separator: xml.boolAttr(elem, "sep"),
            equalWidth: xml.boolAttr(elem, "equalWidth", true),
            columns: xml.elements(elem, "col")
                .map(e => ({
                width: xml.lengthAttr(e, "w"),
                space: xml.lengthAttr(e, "space")
            }))
        };
    }
    function parsePageNumber(elem, xml) {
        return {
            chapSep: xml.attr(elem, "chapSep"),
            chapStyle: xml.attr(elem, "chapStyle"),
            format: xml.attr(elem, "fmt"),
            start: xml.intAttr(elem, "start")
        };
    }
    function parseFooterHeaderReference(elem, xml) {
        return {
            id: xml.attr(elem, "id"),
            type: xml.attr(elem, "type"),
        };
    }

    function parseLineSpacing(elem, xml) {
        return {
            before: xml.lengthAttr(elem, "before"),
            after: xml.lengthAttr(elem, "after"),
            line: xml.intAttr(elem, "line"),
            lineRule: xml.attr(elem, "lineRule")
        };
    }

    function parseRunProperties(elem, xml) {
        let result = {};
        for (let el of xml.elements(elem)) {
            parseRunProperty(el, result, xml);
        }
        return result;
    }
    function parseRunProperty(elem, props, xml) {
        if (parseCommonProperty(elem, props, xml))
            return true;
        return false;
    }

    function parseParagraphProperties(elem, xml) {
        let result = {};
        for (let el of xml.elements(elem)) {
            parseParagraphProperty(el, result, xml);
        }
        return result;
    }
    function parseParagraphProperty(elem, props, xml) {
        if (elem.namespaceURI != ns$1.wordml)
            return false;
        if (parseCommonProperty(elem, props, xml))
            return true;
        switch (elem.localName) {
            case "tabs":
                props.tabs = parseTabs(elem, xml);
                break;
            case "sectPr":
                props.sectionProps = parseSectionProperties(elem, xml);
                break;
            case "numPr":
                props.numbering = parseNumbering$1(elem, xml);
                break;
            case "spacing":
                props.lineSpacing = parseLineSpacing(elem, xml);
                return false;
            case "textAlignment":
                props.textAlignment = xml.attr(elem, "val");
                return false;
            case "keepLines":
                props.keepLines = xml.boolAttr(elem, "val", true);
                break;
            case "keepNext":
                props.keepNext = xml.boolAttr(elem, "val", true);
                break;
            case "pageBreakBefore":
                props.pageBreakBefore = xml.boolAttr(elem, "val", true);
                break;
            case "widowControl":
                props.widowControl = xml.boolAttr(elem, "val", true);
                break;
            case "outlineLvl":
                props.outlineLevel = xml.intAttr(elem, "val");
                break;
            case "pStyle":
                props.styleName = xml.attr(elem, "val");
                break;
            case "rPr":
                props.runProps = parseRunProperties(elem, xml);
                break;
            default:
                return false;
        }
        return true;
    }
    function parseTabs(elem, xml) {
        return xml.elements(elem, "tab")
            .map(e => ({
            position: xml.lengthAttr(e, "pos"),
            leader: xml.attr(e, "leader"),
            style: xml.attr(e, "val")
        }));
    }
    function parseNumbering$1(elem, xml) {
        var result = {};
        for (let e of xml.elements(elem)) {
            switch (e.localName) {
                case "numId":
                    result.id = xml.attr(e, "val");
                    break;
                case "ilvl":
                    result.level = xml.intAttr(e, "val");
                    break;
            }
        }
        return result;
    }

    function parseNumberingPart(elem, xml) {
        let result = {
            numberings: [],
            abstractNumberings: [],
            bulletPictures: []
        };
        for (let e of xml.elements(elem)) {
            switch (e.localName) {
                case "num":
                    result.numberings.push(parseNumbering(e, xml));
                    break;
                case "abstractNum":
                    result.abstractNumberings.push(parseAbstractNumbering(e, xml));
                    break;
                case "numPicBullet":
                    result.bulletPictures.push(parseNumberingBulletPicture(e, xml));
                    break;
            }
        }
        return result;
    }
    function parseNumbering(elem, xml) {
        let result = {
            id: xml.attr(elem, 'numId'),
            overrides: []
        };
        for (let e of xml.elements(elem)) {
            switch (e.localName) {
                case "abstractNumId":
                    result.abstractId = xml.attr(e, "val");
                    break;
                case "lvlOverride":
                    result.overrides.push(parseNumberingLevelOverrride(e, xml));
                    break;
            }
        }
        return result;
    }
    function parseAbstractNumbering(elem, xml) {
        let result = {
            id: xml.attr(elem, 'abstractNumId'),
            levels: []
        };
        for (let e of xml.elements(elem)) {
            switch (e.localName) {
                case "name":
                    result.name = xml.attr(e, "val");
                    break;
                case "multiLevelType":
                    result.multiLevelType = xml.attr(e, "val");
                    break;
                case "numStyleLink":
                    result.numberingStyleLink = xml.attr(e, "val");
                    break;
                case "styleLink":
                    result.styleLink = xml.attr(e, "val");
                    break;
                case "lvl":
                    result.levels.push(parseNumberingLevel(e, xml));
                    break;
            }
        }
        return result;
    }
    function parseNumberingLevel(elem, xml) {
        let result = {
            level: xml.intAttr(elem, 'ilvl')
        };
        for (let e of xml.elements(elem)) {
            switch (e.localName) {
                case "start":
                    result.start = xml.attr(e, "val");
                    break;
                case "lvlRestart":
                    result.restart = xml.intAttr(e, "val");
                    break;
                case "numFmt":
                    result.format = xml.attr(e, "val");
                    break;
                case "lvlText":
                    result.text = xml.attr(e, "val");
                    break;
                case "lvlJc":
                    result.justification = xml.attr(e, "val");
                    break;
                case "lvlPicBulletId":
                    result.bulletPictureId = xml.attr(e, "val");
                    break;
                case "pStyle":
                    result.paragraphStyle = xml.attr(e, "val");
                    break;
                case "pPr":
                    result.paragraphProps = parseParagraphProperties(e, xml);
                    break;
                case "rPr":
                    result.runProps = parseRunProperties(e, xml);
                    break;
            }
        }
        return result;
    }
    function parseNumberingLevelOverrride(elem, xml) {
        let result = {
            level: xml.intAttr(elem, 'ilvl')
        };
        for (let e of xml.elements(elem)) {
            switch (e.localName) {
                case "startOverride":
                    result.start = xml.intAttr(e, "val");
                    break;
                case "lvl":
                    result.numberingLevel = parseNumberingLevel(e, xml);
                    break;
            }
        }
        return result;
    }
    function parseNumberingBulletPicture(elem, xml) {
        var pict = xml.element(elem, "pict");
        var shape = pict && xml.element(pict, "shape");
        var imagedata = shape && xml.element(shape, "imagedata");
        return imagedata ? {
            id: xml.attr(elem, "numPicBulletId"),
            referenceId: xml.attr(imagedata, "id"),
            style: xml.attr(shape, "style")
        } : null;
    }

    class NumberingPart extends Part {
        constructor(pkg, path, parser) {
            super(pkg, path);
            this._documentParser = parser;
        }
        parseXml(root) {
            Object.assign(this, parseNumberingPart(root, this._package.xmlParser));
            this.domNumberings = this._documentParser.parseNumberingFile(root);
        }
    }

    class StylesPart extends Part {
        constructor(pkg, path, parser) {
            super(pkg, path);
            this._documentParser = parser;
        }
        parseXml(root) {
            this.styles = this._documentParser.parseStylesFile(root);
        }
    }

    var DomType;
    (function (DomType) {
        DomType["Document"] = "document";
        DomType["Paragraph"] = "paragraph";
        DomType["Run"] = "run";
        DomType["Break"] = "break";
        DomType["NoBreakHyphen"] = "noBreakHyphen";
        DomType["Table"] = "table";
        DomType["Row"] = "row";
        DomType["Cell"] = "cell";
        DomType["Hyperlink"] = "hyperlink";
        DomType["SmartTag"] = "smartTag";
        DomType["Drawing"] = "drawing";
        DomType["Image"] = "image";
        DomType["DrawingShape"] = "drawingShape";
        DomType["DrawingGroup"] = "drawingGroup";
        DomType["Text"] = "text";
        DomType["Tab"] = "tab";
        DomType["Symbol"] = "symbol";
        DomType["BookmarkStart"] = "bookmarkStart";
        DomType["BookmarkEnd"] = "bookmarkEnd";
        DomType["Footer"] = "footer";
        DomType["Header"] = "header";
        DomType["FootnoteReference"] = "footnoteReference";
        DomType["EndnoteReference"] = "endnoteReference";
        DomType["Footnote"] = "footnote";
        DomType["Endnote"] = "endnote";
        DomType["SimpleField"] = "simpleField";
        DomType["ComplexField"] = "complexField";
        DomType["Instruction"] = "instruction";
        DomType["VmlPicture"] = "vmlPicture";
        DomType["MmlMath"] = "mmlMath";
        DomType["MmlMathParagraph"] = "mmlMathParagraph";
        DomType["MmlFraction"] = "mmlFraction";
        DomType["MmlFunction"] = "mmlFunction";
        DomType["MmlFunctionName"] = "mmlFunctionName";
        DomType["MmlNumerator"] = "mmlNumerator";
        DomType["MmlDenominator"] = "mmlDenominator";
        DomType["MmlRadical"] = "mmlRadical";
        DomType["MmlBase"] = "mmlBase";
        DomType["MmlDegree"] = "mmlDegree";
        DomType["MmlSuperscript"] = "mmlSuperscript";
        DomType["MmlSubscript"] = "mmlSubscript";
        DomType["MmlPreSubSuper"] = "mmlPreSubSuper";
        DomType["MmlSubArgument"] = "mmlSubArgument";
        DomType["MmlSuperArgument"] = "mmlSuperArgument";
        DomType["MmlNary"] = "mmlNary";
        DomType["MmlDelimiter"] = "mmlDelimiter";
        DomType["MmlRun"] = "mmlRun";
        DomType["MmlEquationArray"] = "mmlEquationArray";
        DomType["MmlLimit"] = "mmlLimit";
        DomType["MmlLimitLower"] = "mmlLimitLower";
        DomType["MmlMatrix"] = "mmlMatrix";
        DomType["MmlMatrixRow"] = "mmlMatrixRow";
        DomType["MmlBox"] = "mmlBox";
        DomType["MmlBar"] = "mmlBar";
        DomType["MmlGroupChar"] = "mmlGroupChar";
        DomType["MmlAccent"] = "mmlAccent";
        DomType["MmlBorderBox"] = "mmlBorderBox";
        DomType["MmlSubSuperscript"] = "mmlSubSuperscript";
        DomType["MmlPhantom"] = "mmlPhantom";
        DomType["MmlGroup"] = "mmlGroup";
        DomType["VmlElement"] = "vmlElement";
        DomType["Inserted"] = "inserted";
        DomType["Deleted"] = "deleted";
        DomType["DeletedText"] = "deletedText";
        DomType["MoveFrom"] = "moveFrom";
        DomType["MoveTo"] = "moveTo";
        DomType["Comment"] = "comment";
        DomType["CommentReference"] = "commentReference";
        DomType["CommentRangeStart"] = "commentRangeStart";
        DomType["CommentRangeEnd"] = "commentRangeEnd";
        DomType["AltChunk"] = "altChunk";
        DomType["Sdt"] = "sdt";
        DomType["Ruby"] = "ruby";
        DomType["RubyBase"] = "rubyBase";
        DomType["RubyText"] = "rubyText";
        DomType["FitText"] = "fitText";
        DomType["BidiOverride"] = "bidiOverride";
        DomType["Chart"] = "chart";
        DomType["ChartEx"] = "chartEx";
        DomType["SmartArt"] = "smartArt";
        DomType["OleObject"] = "oleObject";
    })(DomType || (DomType = {}));
    class OpenXmlElementBase {
        constructor() {
            this.children = [];
            this.cssStyle = {};
        }
    }

    class WmlHeader extends OpenXmlElementBase {
        constructor() {
            super(...arguments);
            this.type = DomType.Header;
        }
    }
    class WmlFooter extends OpenXmlElementBase {
        constructor() {
            super(...arguments);
            this.type = DomType.Footer;
        }
    }

    class BaseHeaderFooterPart extends Part {
        constructor(pkg, path, parser) {
            super(pkg, path);
            this._documentParser = parser;
        }
        parseXml(root) {
            this.rootElement = this.createRootElement();
            this.rootElement.children = this._documentParser.parseBodyElements(root);
        }
    }
    class HeaderPart extends BaseHeaderFooterPart {
        createRootElement() {
            return new WmlHeader();
        }
    }
    class FooterPart extends BaseHeaderFooterPart {
        createRootElement() {
            return new WmlFooter();
        }
    }

    function parseExtendedProps(root, xmlParser) {
        const result = {};
        for (let el of xmlParser.elements(root)) {
            switch (el.localName) {
                case "Template":
                    result.template = el.textContent;
                    break;
                case "Pages":
                    result.pages = safeParseToInt(el.textContent);
                    break;
                case "Words":
                    result.words = safeParseToInt(el.textContent);
                    break;
                case "Characters":
                    result.characters = safeParseToInt(el.textContent);
                    break;
                case "Application":
                    result.application = el.textContent;
                    break;
                case "Lines":
                    result.lines = safeParseToInt(el.textContent);
                    break;
                case "Paragraphs":
                    result.paragraphs = safeParseToInt(el.textContent);
                    break;
                case "Company":
                    result.company = el.textContent;
                    break;
                case "AppVersion":
                    result.appVersion = el.textContent;
                    break;
            }
        }
        return result;
    }
    function safeParseToInt(value) {
        if (typeof value === 'undefined')
            return;
        return parseInt(value);
    }

    class ExtendedPropsPart extends Part {
        parseXml(root) {
            this.props = parseExtendedProps(root, this._package.xmlParser);
        }
    }

    function parseCoreProps(root, xmlParser) {
        const result = {};
        for (let el of xmlParser.elements(root)) {
            switch (el.localName) {
                case "title":
                    result.title = el.textContent;
                    break;
                case "description":
                    result.description = el.textContent;
                    break;
                case "subject":
                    result.subject = el.textContent;
                    break;
                case "creator":
                    result.creator = el.textContent;
                    break;
                case "keywords":
                    result.keywords = el.textContent;
                    break;
                case "language":
                    result.language = el.textContent;
                    break;
                case "lastModifiedBy":
                    result.lastModifiedBy = el.textContent;
                    break;
                case "revision":
                    el.textContent && (result.revision = parseInt(el.textContent));
                    break;
                case "created":
                    result.created = el.textContent;
                    break;
                case "modified":
                    result.modified = el.textContent;
                    break;
            }
        }
        return result;
    }

    class CorePropsPart extends Part {
        parseXml(root) {
            this.props = parseCoreProps(root, this._package.xmlParser);
        }
    }

    class DmlTheme {
    }
    function parseTheme(elem, xml) {
        var result = new DmlTheme();
        var themeElements = xml.element(elem, "themeElements");
        for (let el of xml.elements(themeElements)) {
            switch (el.localName) {
                case "clrScheme":
                    result.colorScheme = parseColorScheme(el, xml);
                    break;
                case "fontScheme":
                    result.fontScheme = parseFontScheme(el, xml);
                    break;
            }
        }
        return result;
    }
    function parseColorScheme(elem, xml) {
        var result = {
            name: xml.attr(elem, "name"),
            colors: {}
        };
        for (let el of xml.elements(elem)) {
            var srgbClr = xml.element(el, "srgbClr");
            var sysClr = xml.element(el, "sysClr");
            if (srgbClr) {
                result.colors[el.localName] = xml.attr(srgbClr, "val");
            }
            else if (sysClr) {
                result.colors[el.localName] = xml.attr(sysClr, "lastClr");
            }
        }
        return result;
    }
    function parseFontScheme(elem, xml) {
        var result = {
            name: xml.attr(elem, "name"),
        };
        for (let el of xml.elements(elem)) {
            switch (el.localName) {
                case "majorFont":
                    result.majorFont = parseFontInfo(el, xml);
                    break;
                case "minorFont":
                    result.minorFont = parseFontInfo(el, xml);
                    break;
            }
        }
        return result;
    }
    function parseFontInfo(elem, xml) {
        return {
            latinTypeface: xml.elementAttr(elem, "latin", "typeface"),
            eaTypeface: xml.elementAttr(elem, "ea", "typeface"),
            csTypeface: xml.elementAttr(elem, "cs", "typeface"),
        };
    }

    class ThemePart extends Part {
        constructor(pkg, path) {
            super(pkg, path);
        }
        parseXml(root) {
            this.theme = parseTheme(root, this._package.xmlParser);
        }
    }

    class WmlBaseNote {
    }
    class WmlFootnote extends WmlBaseNote {
        constructor() {
            super(...arguments);
            this.type = DomType.Footnote;
        }
    }
    class WmlEndnote extends WmlBaseNote {
        constructor() {
            super(...arguments);
            this.type = DomType.Endnote;
        }
    }

    class BaseNotePart extends Part {
        constructor(pkg, path, parser) {
            super(pkg, path);
            this._documentParser = parser;
        }
    }
    class FootnotesPart extends BaseNotePart {
        constructor(pkg, path, parser) {
            super(pkg, path, parser);
        }
        parseXml(root) {
            this.notes = this._documentParser.parseNotes(root, "footnote", WmlFootnote);
        }
    }
    class EndnotesPart extends BaseNotePart {
        constructor(pkg, path, parser) {
            super(pkg, path, parser);
        }
        parseXml(root) {
            this.notes = this._documentParser.parseNotes(root, "endnote", WmlEndnote);
        }
    }

    function parseSettings(elem, xml) {
        var result = {};
        for (let el of xml.elements(elem)) {
            switch (el.localName) {
                case "defaultTabStop":
                    result.defaultTabStop = xml.lengthAttr(el, "val");
                    break;
                case "footnotePr":
                    result.footnoteProps = parseNoteProperties(el, xml);
                    break;
                case "endnotePr":
                    result.endnoteProps = parseNoteProperties(el, xml);
                    break;
                case "autoHyphenation":
                    result.autoHyphenation = xml.boolAttr(el, "val");
                    break;
                case "evenAndOddHeaders":
                    result.evenAndOddHeaders = xml.boolAttr(el, "val", true);
                    break;
                case "documentProtection":
                    result.documentProtection = parseDocumentProtection(el, xml);
                    break;
            }
        }
        return result;
    }
    const ALLOWED_EDIT = new Set(['readOnly', 'trackedChanges', 'comments', 'forms', 'none']);
    function parseDocumentProtection(elem, xml) {
        const editAttr = xml.attr(elem, "edit");
        return {
            edit: ALLOWED_EDIT.has(editAttr) ? editAttr : undefined,
            enforcement: xml.boolAttr(elem, "enforcement", false),
            formatting: xml.boolAttr(elem, "formatting", false),
        };
    }
    function parseNoteProperties(elem, xml) {
        var result = {
            defaultNoteIds: []
        };
        for (let el of xml.elements(elem)) {
            switch (el.localName) {
                case "numFmt":
                    result.nummeringFormat = xml.attr(el, "val");
                    break;
                case "footnote":
                case "endnote":
                    result.defaultNoteIds.push(xml.attr(el, "id"));
                    break;
            }
        }
        return result;
    }

    class SettingsPart extends Part {
        constructor(pkg, path) {
            super(pkg, path);
        }
        parseXml(root) {
            this.settings = parseSettings(root, this._package.xmlParser);
        }
    }

    function parseCustomProps(root, xml) {
        return xml.elements(root, "property").map(e => {
            const firstChild = e.firstChild;
            return {
                formatId: xml.attr(e, "fmtid"),
                name: xml.attr(e, "name"),
                type: firstChild.nodeName,
                value: firstChild.textContent
            };
        });
    }

    class CustomPropsPart extends Part {
        parseXml(root) {
            this.props = parseCustomProps(root, this._package.xmlParser);
        }
    }

    const SAFE_PARA_ID = /^[A-Za-z0-9_-]+$/;
    class CommentsPart extends Part {
        constructor(pkg, path, parser) {
            super(pkg, path);
            this.topLevelComments = [];
            this._documentParser = parser;
        }
        parseXml(root) {
            this.comments = this._documentParser.parseComments(root);
            this.commentMap = keyBy(this.comments, x => x.id);
        }
        buildThreading(extendedComments) {
            if (!extendedComments || extendedComments.length === 0) {
                this.topLevelComments = [...this.comments];
                return;
            }
            const extMap = new Map();
            for (const ext of extendedComments) {
                if (ext.paraId && SAFE_PARA_ID.test(ext.paraId)) {
                    extMap.set(ext.paraId, ext);
                }
            }
            const paraIdToComment = new Map();
            for (const comment of this.comments) {
                if (comment.paraId && SAFE_PARA_ID.test(comment.paraId)) {
                    paraIdToComment.set(comment.paraId, comment);
                    const ext = extMap.get(comment.paraId);
                    if (ext) {
                        comment.done = ext.done;
                    }
                }
            }
            for (const ext of extendedComments) {
                if (ext.paraIdParent && SAFE_PARA_ID.test(ext.paraIdParent) && ext.paraId && SAFE_PARA_ID.test(ext.paraId)) {
                    const child = paraIdToComment.get(ext.paraId);
                    const parent = paraIdToComment.get(ext.paraIdParent);
                    if (child && parent) {
                        child.parentCommentId = parent.id;
                        parent.replies.push(child);
                    }
                }
            }
            this.topLevelComments = this.comments.filter(c => !c.parentCommentId);
        }
    }

    class CommentsExtendedPart extends Part {
        constructor(pkg, path) {
            super(pkg, path);
            this.comments = [];
        }
        parseXml(root) {
            const xml = this._package.xmlParser;
            for (let el of xml.elements(root, "commentEx")) {
                this.comments.push({
                    paraId: xml.attr(el, 'paraId'),
                    paraIdParent: xml.attr(el, 'paraIdParent'),
                    done: xml.boolAttr(el, 'done')
                });
            }
            this.commentMap = keyBy(this.comments, x => x.paraId);
        }
    }

    const DEFAULT_THEME_PALETTE = {
        accent1: '#4472C4',
        accent2: '#ED7D31',
        accent3: '#A5A5A5',
        accent4: '#FFC000',
        accent5: '#5B9BD5',
        accent6: '#70AD47',
        dk1: '#000000',
        lt1: '#FFFFFF',
        dk2: '#44546A',
        lt2: '#E7E6E6',
        bg1: '#FFFFFF',
        bg2: '#E7E6E6',
        tx1: '#000000',
        tx2: '#44546A',
        hlink: '#0563C1',
        folHlink: '#954F72',
    };
    const ALLOWED_SLOTS = new Set(Object.keys(DEFAULT_THEME_PALETTE));
    const DEFAULT_FALLBACK_COLOUR = '#808080';
    function resolveColour(ref, palette) {
        if (!ref)
            return DEFAULT_FALLBACK_COLOUR;
        let base = null;
        if (ref.hex) {
            base = ref.hex;
        }
        else if (ref.scheme && ALLOWED_SLOTS.has(ref.scheme)) {
            const fromTheme = palette && palette[ref.scheme];
            const sanitised = sanitizeCssColor(fromTheme);
            base = sanitised ?? DEFAULT_THEME_PALETTE[ref.scheme];
        }
        if (!base)
            return DEFAULT_FALLBACK_COLOUR;
        const rgb = hexToRgb(base);
        if (!rgb)
            return DEFAULT_FALLBACK_COLOUR;
        const mods = ref.mods;
        if (!mods)
            return rgbToHex(rgb);
        let { r, g, b } = rgb;
        if (typeof mods.lumMod === 'number' || typeof mods.lumOff === 'number') {
            const hsl = rgbToHsl(r, g, b);
            const mul = typeof mods.lumMod === 'number'
                ? clamp(mods.lumMod / 100000, 0, 1)
                : 1;
            const add = typeof mods.lumOff === 'number'
                ? clamp(mods.lumOff / 100000, -1, 1)
                : 0;
            hsl.l = clamp(hsl.l * mul + add, 0, 1);
            const out = hslToRgb(hsl.h, hsl.s, hsl.l);
            r = out.r;
            g = out.g;
            b = out.b;
        }
        if (typeof mods.tint === 'number') {
            const t = clamp(mods.tint / 100000, 0, 1);
            r = Math.round(r + (255 - r) * t);
            g = Math.round(g + (255 - g) * t);
            b = Math.round(b + (255 - b) * t);
        }
        if (typeof mods.shade === 'number') {
            const s = clamp(mods.shade / 100000, 0, 1);
            r = Math.round(r * s);
            g = Math.round(g * s);
            b = Math.round(b * s);
        }
        return rgbToHex({ r, g, b });
    }
    function isAllowedSchemeSlot(val) {
        return typeof val === 'string' && ALLOWED_SLOTS.has(val);
    }
    function resolveSchemeColor(val, palette) {
        if (!isAllowedSchemeSlot(val))
            return null;
        const themed = palette?.[val];
        if (themed)
            return themed;
        return DEFAULT_THEME_PALETTE[val] ?? null;
    }
    function buildThemeColorReference(slot, themeTint, themeShade) {
        if (!isAllowedSchemeSlot(slot))
            return null;
        const parts = [slot];
        if (typeof themeTint === 'string' && /^[0-9a-fA-F]{1,2}$/.test(themeTint)) {
            parts.push(`tint=${themeTint.toLowerCase()}`);
        }
        if (typeof themeShade === 'string' && /^[0-9a-fA-F]{1,2}$/.test(themeShade)) {
            parts.push(`shade=${themeShade.toLowerCase()}`);
        }
        return parts.join(':');
    }
    function parseThemeColorReference(raw) {
        if (typeof raw !== 'string')
            return null;
        const trimmed = raw.trim();
        if (!trimmed || trimmed === 'auto')
            return null;
        const hexSan = sanitizeCssColor(trimmed);
        if (hexSan && /^#?[0-9a-fA-F]{3}([0-9a-fA-F]{3})?$/.test(trimmed)) {
            return { hex: hexSan };
        }
        const segments = trimmed.split(/[:\s]+/).filter(Boolean);
        if (segments.length === 0)
            return null;
        const slot = segments[0];
        if (!isAllowedSchemeSlot(slot))
            return null;
        const ref = { scheme: slot };
        const mods = {};
        let gotMod = false;
        for (let i = 1; i < segments.length; i++) {
            const eq = segments[i].indexOf('=');
            if (eq < 0)
                continue;
            const key = segments[i].slice(0, eq);
            const val = segments[i].slice(eq + 1);
            if (!/^[0-9a-fA-F]{1,2}$/.test(val))
                continue;
            const byte = parseInt(val, 16);
            if (!Number.isFinite(byte))
                continue;
            const scaled = Math.round((byte / 255) * 100000);
            if (key === 'tint') {
                mods.tint = scaled;
                gotMod = true;
            }
            else if (key === 'shade') {
                mods.shade = scaled;
                gotMod = true;
            }
        }
        if (gotMod)
            ref.mods = mods;
        return ref;
    }
    function clamp(v, lo, hi) {
        if (v < lo)
            return lo;
        if (v > hi)
            return hi;
        return v;
    }
    function hexToRgb(hex) {
        if (typeof hex !== 'string')
            return null;
        let h = hex.trim();
        if (h.startsWith('#'))
            h = h.slice(1);
        if (h.length === 3)
            h = h.split('').map((c) => c + c).join('');
        if (h.length !== 6)
            return null;
        if (!/^[0-9a-fA-F]{6}$/.test(h))
            return null;
        return {
            r: parseInt(h.slice(0, 2), 16),
            g: parseInt(h.slice(2, 4), 16),
            b: parseInt(h.slice(4, 6), 16),
        };
    }
    function rgbToHex(rgb) {
        const to2 = (n) => clamp(Math.round(n), 0, 255).toString(16).padStart(2, '0');
        return `#${to2(rgb.r)}${to2(rgb.g)}${to2(rgb.b)}`;
    }
    function rgbToHsl(r, g, b) {
        const rn = r / 255, gn = g / 255, bn = b / 255;
        const max = Math.max(rn, gn, bn);
        const min = Math.min(rn, gn, bn);
        let h = 0, s = 0;
        const l = (max + min) / 2;
        if (max !== min) {
            const d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            switch (max) {
                case rn:
                    h = (gn - bn) / d + (gn < bn ? 6 : 0);
                    break;
                case gn:
                    h = (bn - rn) / d + 2;
                    break;
                default: h = (rn - gn) / d + 4;
            }
            h /= 6;
        }
        return { h, s, l };
    }
    function hslToRgb(h, s, l) {
        if (s === 0) {
            const v = Math.round(l * 255);
            return { r: v, g: v, b: v };
        }
        const hue2rgb = (p, q, t) => {
            let tt = t;
            if (tt < 0)
                tt += 1;
            if (tt > 1)
                tt -= 1;
            if (tt < 1 / 6)
                return p + (q - p) * 6 * tt;
            if (tt < 1 / 2)
                return q;
            if (tt < 2 / 3)
                return p + (q - p) * (2 / 3 - tt) * 6;
            return p;
        };
        const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        const p = 2 * l - q;
        return {
            r: Math.round(hue2rgb(p, q, h + 1 / 3) * 255),
            g: Math.round(hue2rgb(p, q, h) * 255),
            b: Math.round(hue2rgb(p, q, h - 1 / 3) * 255),
        };
    }

    class ChartPart extends Part {
        constructor(pkg, path) {
            super(pkg, path);
        }
        parseXml(root) {
            this.chart = parseChartSpace(root, this.path);
        }
    }
    function parseChartSpace(chartSpace, path) {
        const key = deriveKey$1(path);
        const chart = globalXmlParser.element(chartSpace, "chart");
        const emptyAxis = { line: null, tickLabel: null, gridline: null };
        if (!chart) {
            return {
                key, title: "", showLegend: false, kind: "unknown",
                grouping: "clustered", series: [],
                catAxis: { ...emptyAxis }, valAxis: { ...emptyAxis },
            };
        }
        const titleEl = globalXmlParser.element(chart, "title");
        const legendEl = globalXmlParser.element(chart, "legend");
        const plotArea = globalXmlParser.element(chart, "plotArea");
        const title = titleEl ? extractRichText(titleEl) : "";
        const showLegend = legendEl != null;
        let kind = "unknown";
        let grouping = "clustered";
        let series = [];
        let catAxis = { ...emptyAxis };
        let valAxis = { ...emptyAxis };
        if (plotArea) {
            for (const child of globalXmlParser.elements(plotArea)) {
                const k = localNameToKind(child.localName);
                if (k === "unknown")
                    continue;
                kind = k;
                grouping = globalXmlParser.attr(globalXmlParser.element(child, "grouping") ?? null, "val") ?? "clustered";
                if (k === "column" || k === "bar") {
                    const barDir = globalXmlParser.attr(globalXmlParser.element(child, "barDir") ?? null, "val");
                    if (barDir === "bar")
                        kind = "bar";
                    else if (barDir === "col")
                        kind = "column";
                }
                series = globalXmlParser.elements(child, "ser").map(parseSeries);
                break;
            }
            const catAxEl = globalXmlParser.element(plotArea, "catAx");
            if (catAxEl)
                catAxis = parseAxisStyle(catAxEl);
            const valAxEl = globalXmlParser.element(plotArea, "valAx");
            if (valAxEl)
                valAxis = parseAxisStyle(valAxEl);
        }
        return { key, title, showLegend, kind, grouping, series, catAxis, valAxis };
    }
    function parseAxisStyle(axEl) {
        const style = { line: null, tickLabel: null, gridline: null };
        const axSpPr = globalXmlParser.element(axEl, "spPr");
        if (axSpPr) {
            const ln = globalXmlParser.element(axSpPr, "ln");
            if (ln) {
                style.line = parseSolidFillRef(ln);
            }
            if (style.line == null) {
                style.line = parseSolidFillRef(axSpPr);
            }
        }
        const txPr = globalXmlParser.element(axEl, "txPr");
        if (txPr) {
            const p = globalXmlParser.element(txPr, "p");
            const pPr = p ? globalXmlParser.element(p, "pPr") : null;
            const defRPr = pPr ? globalXmlParser.element(pPr, "defRPr") : null;
            if (defRPr) {
                style.tickLabel = parseSolidFillRef(defRPr);
            }
        }
        const grid = globalXmlParser.element(axEl, "majorGridlines");
        if (grid) {
            const gSpPr = globalXmlParser.element(grid, "spPr");
            if (gSpPr) {
                const ln = globalXmlParser.element(gSpPr, "ln");
                if (ln) {
                    style.gridline = parseSolidFillRef(ln);
                }
                if (style.gridline == null) {
                    style.gridline = parseSolidFillRef(gSpPr);
                }
            }
        }
        return style;
    }
    function parseSolidFillRef(spPr) {
        const solidFill = globalXmlParser.element(spPr, "solidFill");
        if (!solidFill)
            return null;
        const srgb = globalXmlParser.element(solidFill, "srgbClr");
        if (srgb) {
            const raw = globalXmlParser.attr(srgb, "val");
            const safe = sanitizeCssColor(raw);
            return safe ? { kind: "literal", color: safe } : null;
        }
        const scheme = globalXmlParser.element(solidFill, "schemeClr");
        if (scheme) {
            const raw = globalXmlParser.attr(scheme, "val");
            if (isAllowedSchemeSlot(raw))
                return { kind: "scheme", slot: raw };
        }
        return null;
    }
    function deriveKey$1(path) {
        const segs = path.split("/");
        const file = segs[segs.length - 1] ?? "";
        const dot = file.lastIndexOf(".");
        return dot > 0 ? file.slice(0, dot) : file;
    }
    function localNameToKind(name) {
        switch (name) {
            case "barChart": return "column";
            case "bar3DChart": return "column";
            case "lineChart": return "line";
            case "line3DChart": return "line";
            case "pieChart": return "pie";
            case "pie3DChart": return "pie";
            case "doughnutChart": return "pie";
            default: return "unknown";
        }
    }
    function extractRichText(titleEl) {
        const strRef = findDescendant(titleEl, "strRef");
        if (strRef) {
            const cache = globalXmlParser.element(strRef, "strCache");
            if (cache) {
                const pt = globalXmlParser.element(cache, "pt");
                if (pt) {
                    const v = globalXmlParser.element(pt, "v");
                    if (v)
                        return v.textContent ?? "";
                }
            }
        }
        const parts = [];
        walkTextRuns$1(titleEl, parts);
        return parts.join("");
    }
    function walkTextRuns$1(node, out) {
        for (const c of globalXmlParser.elements(node)) {
            if (c.localName === "t") {
                out.push(c.textContent ?? "");
            }
            else {
                walkTextRuns$1(c, out);
            }
        }
    }
    function findDescendant(node, localName) {
        for (const c of globalXmlParser.elements(node)) {
            if (c.localName === localName)
                return c;
            const d = findDescendant(c, localName);
            if (d)
                return d;
        }
        return null;
    }
    function parseSeries(serEl) {
        const title = parseSeriesTitle(serEl);
        const color = parseSeriesColor(serEl);
        const categories = parseCategoryLabels(globalXmlParser.element(serEl, "cat"));
        const values = parseNumericValues(globalXmlParser.element(serEl, "val"));
        const dataPointOverrides = parseDataPointOverrides$1(serEl);
        return { title, color, categories, values, dataPointOverrides };
    }
    function parseDataPointOverrides$1(serEl) {
        const out = new Map();
        for (const dPt of globalXmlParser.elements(serEl, "dPt")) {
            const idxEl = globalXmlParser.element(dPt, "idx");
            const rawIdx = idxEl ? globalXmlParser.attr(idxEl, "val") : null;
            const idx = rawIdx != null ? parseInt(rawIdx, 10) : NaN;
            if (!Number.isFinite(idx) || idx < 0 || idx >= MAX_POINTS$1)
                continue;
            const spPr = globalXmlParser.element(dPt, "spPr");
            if (!spPr)
                continue;
            const color = parseSolidFillColor$1(spPr);
            if (color == null)
                continue;
            out.set(idx, { color });
        }
        return out;
    }
    function parseSeriesTitle(serEl) {
        const tx = globalXmlParser.element(serEl, "tx");
        if (!tx)
            return "";
        const strRef = globalXmlParser.element(tx, "strRef");
        if (strRef) {
            const cache = globalXmlParser.element(strRef, "strCache");
            if (cache) {
                const pt = globalXmlParser.element(cache, "pt");
                if (pt) {
                    const v = globalXmlParser.element(pt, "v");
                    if (v)
                        return v.textContent ?? "";
                }
            }
        }
        const v = globalXmlParser.element(tx, "v");
        if (v)
            return v.textContent ?? "";
        const rich = globalXmlParser.element(tx, "rich");
        if (rich) {
            const parts = [];
            walkTextRuns$1(rich, parts);
            return parts.join("");
        }
        return "";
    }
    function parseSeriesColor(serEl) {
        const spPr = globalXmlParser.element(serEl, "spPr");
        if (!spPr)
            return null;
        return parseSolidFillColor$1(spPr);
    }
    function parseSolidFillColor$1(spPr) {
        const solidFill = globalXmlParser.element(spPr, "solidFill");
        if (!solidFill)
            return null;
        const srgb = globalXmlParser.element(solidFill, "srgbClr");
        if (srgb) {
            const raw = globalXmlParser.attr(srgb, "val");
            return sanitizeCssColor(raw);
        }
        const scheme = globalXmlParser.element(solidFill, "schemeClr");
        if (scheme) {
            const raw = globalXmlParser.attr(scheme, "val");
            const resolved = resolveSchemeColor(raw);
            return resolved ? sanitizeCssColor(resolved) : null;
        }
        return null;
    }
    function parseCategoryLabels(catEl) {
        if (!catEl)
            return [];
        const cache = findCache(catEl, ["strCache", "numCache"]);
        if (!cache)
            return [];
        return extractPoints(cache, (v) => v);
    }
    function parseNumericValues(valEl) {
        if (!valEl)
            return [];
        const cache = findCache(valEl, ["numCache", "strCache"]);
        if (!cache)
            return [];
        return extractPoints(cache, (v) => {
            const n = parseFloat(v);
            return Number.isFinite(n) ? n : NaN;
        });
    }
    function findCache(parent, localNames) {
        for (const name of localNames) {
            const ref = globalXmlParser.element(parent, name === "numCache" ? "numRef" : "strRef");
            if (ref) {
                const cache = globalXmlParser.element(ref, name);
                if (cache)
                    return cache;
            }
        }
        for (const name of localNames) {
            const cache = globalXmlParser.element(parent, name);
            if (cache)
                return cache;
        }
        return null;
    }
    const MAX_POINTS$1 = 4096;
    function extractPoints(cache, mapValue) {
        const pts = globalXmlParser.elements(cache, "pt");
        const pairs = [];
        for (const pt of pts) {
            const rawIdx = globalXmlParser.attr(pt, "idx");
            const idx = rawIdx != null ? parseInt(rawIdx, 10) : NaN;
            if (!Number.isFinite(idx) || idx < 0 || idx >= MAX_POINTS$1)
                continue;
            const v = globalXmlParser.element(pt, "v");
            const text = v ? (v.textContent ?? "") : "";
            pairs.push({ idx, value: mapValue(text) });
        }
        pairs.sort((a, b) => a.idx - b.idx);
        if (pairs.length === 0)
            return [];
        const maxIdx = pairs[pairs.length - 1].idx;
        const out = new Array(maxIdx + 1);
        for (const { idx, value } of pairs)
            out[idx] = value;
        return out;
    }

    const CHARTEX_KINDS = {
        sunburst: "sunburst",
        waterfall: "waterfall",
        funnel: "funnel",
        treemap: "treemap",
        histogram: "histogram",
        pareto: "pareto",
        boxWhisker: "box_whisker",
        clusteredColumn: "unknown",
        regionMap: "unknown",
    };
    const MAX_POINTS = 4096;
    class ChartExPart extends Part {
        constructor(pkg, path) {
            super(pkg, path);
        }
        parseXml(root) {
            this.chart = parseChartExSpace(root, this.path);
        }
    }
    function parseChartExSpace(root, path) {
        const key = deriveKey(path);
        const chart = globalXmlParser.element(root, "chart");
        const plotArea = chart ? globalXmlParser.element(chart, "plotArea") : null;
        const plotSurface = plotArea ? globalXmlParser.element(plotArea, "plotSurface") : null;
        const plotAreaRegion = plotArea ? globalXmlParser.element(plotArea, "plotAreaRegion") : null;
        let kind = "unknown";
        let firstSeries = null;
        if (plotAreaRegion) {
            for (const seriesEl of globalXmlParser.elements(plotAreaRegion, "series")) {
                const layoutId = globalXmlParser.attr(seriesEl, "layoutId");
                if (layoutId && CHARTEX_KINDS[layoutId]) {
                    kind = CHARTEX_KINDS[layoutId];
                    firstSeries = seriesEl;
                    break;
                }
            }
        }
        if (kind === "unknown" && plotSurface) {
            const layoutId = globalXmlParser.attr(plotSurface, "layoutId");
            if (layoutId && CHARTEX_KINDS[layoutId]) {
                kind = CHARTEX_KINDS[layoutId];
            }
        }
        const title = chart ? extractTitle(chart) : "";
        if ((kind === "sunburst" || kind === "treemap") && firstSeries) {
            const dataModel = tryParseTreeModel(root, firstSeries, key, title, kind);
            if (dataModel)
                return dataModel;
        }
        else if (kind === "waterfall" && firstSeries) {
            const dataModel = tryParseWaterfallModel(root, firstSeries, key, title);
            if (dataModel)
                return dataModel;
        }
        else if (kind === "funnel" && firstSeries) {
            const dataModel = tryParseFunnelModel(root, firstSeries, key, title);
            if (dataModel)
                return dataModel;
        }
        else if (kind === "histogram" && firstSeries) {
            const dataModel = tryParseHistogramModel(root, firstSeries, key, title);
            if (dataModel)
                return dataModel;
        }
        const placeholder = { shape: "placeholder", key, title, kind };
        return placeholder;
    }
    function tryParseTreeModel(root, seriesEl, key, title, kind) {
        const dataEl = findDataBlock(root, seriesEl);
        if (!dataEl)
            return null;
        const catDim = findDimension(dataEl, "strDim", "cat")
            ?? findDimension(dataEl, "numDim", "cat");
        const valDim = findDimension(dataEl, "numDim", "val");
        if (!catDim || !valDim)
            return null;
        const catLevels = parseStringLevels(catDim);
        const values = parseNumericLevel(valDim);
        if (catLevels.length === 0 || values.length === 0)
            return null;
        const dataPointColors = parseDataPointOverrides(seriesEl);
        const root0 = buildCategoryTree(catLevels, values, dataPointColors);
        if (!root0 || root0.children.length === 0)
            return null;
        const maxDepth = computeMaxDepth(root0);
        return { shape: "data", key, title, kind, root: root0, maxDepth };
    }
    function findDataBlock(root, seriesEl) {
        const dataIdEl = globalXmlParser.element(seriesEl, "dataId");
        const dataId = dataIdEl ? globalXmlParser.attr(dataIdEl, "val") : null;
        const chartData = globalXmlParser.element(root, "chartData");
        if (!chartData)
            return null;
        for (const d of globalXmlParser.elements(chartData, "data")) {
            if (dataId == null || globalXmlParser.attr(d, "id") === dataId) {
                return d;
            }
        }
        return null;
    }
    function parseFlatDimensions(dataEl) {
        const catDim = findDimension(dataEl, "strDim", "cat");
        const valDim = findDimension(dataEl, "numDim", "val");
        if (!valDim)
            return null;
        const values = parseNumericLevel(valDim);
        if (values.length === 0)
            return null;
        let labels;
        if (catDim) {
            const levels = parseStringLevels(catDim);
            labels = levels.length > 0 ? levels[0].points : [];
        }
        else {
            labels = [];
        }
        while (labels.length < values.length)
            labels.push("");
        if (labels.length > values.length)
            labels.length = values.length;
        return { labels, values };
    }
    function tryParseWaterfallModel(root, seriesEl, key, title) {
        const dataEl = findDataBlock(root, seriesEl);
        if (!dataEl)
            return null;
        const flat = parseFlatDimensions(dataEl);
        if (!flat)
            return null;
        const subtotalSet = new Set();
        const layoutPr = globalXmlParser.element(seriesEl, "layoutPr");
        if (layoutPr) {
            const subtotals = globalXmlParser.element(layoutPr, "subtotals");
            if (subtotals) {
                for (const st of globalXmlParser.elements(subtotals, "subtotal")) {
                    const rawIdx = globalXmlParser.attr(st, "idx");
                    const idx = rawIdx != null ? parseInt(rawIdx, 10) : NaN;
                    if (Number.isFinite(idx) && idx >= 0 && idx < MAX_POINTS) {
                        subtotalSet.add(idx);
                    }
                }
            }
        }
        const overrides = parseDataPointOverrides(seriesEl);
        const points = flat.values.map((v, i) => {
            const value = Number.isFinite(v) ? v : 0;
            let type = "normal";
            if (subtotalSet.has(i)) {
                type = i === flat.values.length - 1 ? "total" : "subtotal";
            }
            return {
                label: flat.labels[i] ?? "",
                value,
                type,
                color: overrides.get(i) ?? null,
            };
        });
        if (points.length === 0)
            return null;
        return { shape: "data", key, title, kind: "waterfall", points };
    }
    function tryParseFunnelModel(root, seriesEl, key, title) {
        const dataEl = findDataBlock(root, seriesEl);
        if (!dataEl)
            return null;
        const flat = parseFlatDimensions(dataEl);
        if (!flat)
            return null;
        const overrides = parseDataPointOverrides(seriesEl);
        const points = flat.values.map((v, i) => ({
            label: flat.labels[i] ?? "",
            value: Number.isFinite(v) && v >= 0 ? v : 0,
            color: overrides.get(i) ?? null,
        }));
        if (points.length === 0)
            return null;
        return { shape: "data", key, title, kind: "funnel", points };
    }
    function tryParseHistogramModel(root, seriesEl, key, title) {
        const dataEl = findDataBlock(root, seriesEl);
        if (!dataEl)
            return null;
        const valDim = findDimension(dataEl, "numDim", "val");
        if (!valDim)
            return null;
        const rawValues = parseNumericLevel(valDim);
        const values = rawValues.filter((v) => Number.isFinite(v));
        if (values.length === 0)
            return null;
        const binning = parseBinning(seriesEl);
        const seriesSpPr = globalXmlParser.element(seriesEl, "spPr");
        const seriesColor = seriesSpPr ? parseSolidFillColor(seriesSpPr) : null;
        const dataPointOverrides = parseDataPointOverrides(seriesEl);
        return {
            shape: "data",
            key,
            title,
            kind: "histogram",
            values,
            binning,
            seriesColor,
            dataPointOverrides,
        };
    }
    function parseBinning(seriesEl) {
        const layoutPr = globalXmlParser.element(seriesEl, "layoutPr");
        const out = {
            binSize: null, binCount: null, underflow: null, overflow: null,
        };
        if (!layoutPr)
            return out;
        const binning = globalXmlParser.element(layoutPr, "binning");
        if (!binning)
            return out;
        const parseFiniteAttr = (name) => {
            const raw = globalXmlParser.attr(binning, name);
            if (raw == null)
                return null;
            const n = parseFloat(raw);
            return Number.isFinite(n) ? n : null;
        };
        out.binSize = parseFiniteAttr("binSize");
        const binCountRaw = parseFiniteAttr("binCount");
        out.binCount = binCountRaw != null && binCountRaw >= 1 && binCountRaw <= MAX_POINTS
            ? Math.floor(binCountRaw)
            : null;
        if (out.binSize != null && !(out.binSize > 0))
            out.binSize = null;
        out.underflow = parseFiniteAttr("underflow");
        out.overflow = parseFiniteAttr("overflow");
        return out;
    }
    function findDimension(dataEl, localName, type) {
        for (const d of globalXmlParser.elements(dataEl, localName)) {
            if (globalXmlParser.attr(d, "type") === type)
                return d;
        }
        return null;
    }
    function parseStringLevels(dim) {
        const out = [];
        for (const lvl of globalXmlParser.elements(dim, "lvl")) {
            const level = parseStringLevel(lvl);
            if (level)
                out.push(level);
        }
        return out;
    }
    function parseStringLevel(lvl) {
        const ptCountAttr = globalXmlParser.attr(lvl, "ptCount");
        const ptCount = ptCountAttr != null ? parseInt(ptCountAttr, 10) : NaN;
        const declared = Number.isFinite(ptCount) && ptCount >= 0 && ptCount < MAX_POINTS
            ? ptCount : 0;
        const points = new Array(declared).fill("");
        const parents = new Array(declared).fill(-1);
        let maxIdx = declared - 1;
        for (const pt of globalXmlParser.elements(lvl, "pt")) {
            const rawIdx = globalXmlParser.attr(pt, "idx");
            const idx = rawIdx != null ? parseInt(rawIdx, 10) : NaN;
            if (!Number.isFinite(idx) || idx < 0 || idx >= MAX_POINTS)
                continue;
            if (idx > maxIdx) {
                while (points.length <= idx)
                    points.push("");
                while (parents.length <= idx)
                    parents.push(-1);
                maxIdx = idx;
            }
            points[idx] = pt.textContent ?? "";
            const parentAttr = globalXmlParser.attr(pt, "parent") ?? globalXmlParser.attr(pt, "parentIdx");
            if (parentAttr != null) {
                const p = parseInt(parentAttr, 10);
                if (Number.isFinite(p) && p >= 0 && p < MAX_POINTS)
                    parents[idx] = p;
            }
        }
        if (points.length === 0)
            return null;
        return { points, parents };
    }
    function parseNumericLevel(dim) {
        const lvl = globalXmlParser.element(dim, "lvl");
        if (!lvl)
            return [];
        const ptCountAttr = globalXmlParser.attr(lvl, "ptCount");
        const ptCount = ptCountAttr != null ? parseInt(ptCountAttr, 10) : NaN;
        const declared = Number.isFinite(ptCount) && ptCount >= 0 && ptCount < MAX_POINTS
            ? ptCount : 0;
        const values = new Array(declared).fill(NaN);
        let maxIdx = declared - 1;
        for (const pt of globalXmlParser.elements(lvl, "pt")) {
            const rawIdx = globalXmlParser.attr(pt, "idx");
            const idx = rawIdx != null ? parseInt(rawIdx, 10) : NaN;
            if (!Number.isFinite(idx) || idx < 0 || idx >= MAX_POINTS)
                continue;
            if (idx > maxIdx) {
                while (values.length <= idx)
                    values.push(NaN);
                maxIdx = idx;
            }
            const raw = pt.textContent ?? "";
            const n = parseFloat(raw);
            values[idx] = Number.isFinite(n) ? n : NaN;
        }
        return values;
    }
    function parseDataPointOverrides(seriesEl) {
        const out = new Map();
        for (const dPt of globalXmlParser.elements(seriesEl, "dataPt")) {
            const rawIdx = globalXmlParser.attr(dPt, "idx");
            const idx = rawIdx != null ? parseInt(rawIdx, 10) : NaN;
            if (!Number.isFinite(idx) || idx < 0 || idx >= MAX_POINTS)
                continue;
            const spPr = globalXmlParser.element(dPt, "spPr");
            if (!spPr)
                continue;
            const color = parseSolidFillColor(spPr);
            if (color)
                out.set(idx, color);
        }
        return out;
    }
    function parseSolidFillColor(spPr) {
        const solidFill = globalXmlParser.element(spPr, "solidFill");
        if (!solidFill)
            return null;
        const srgb = globalXmlParser.element(solidFill, "srgbClr");
        if (srgb) {
            const raw = globalXmlParser.attr(srgb, "val");
            return sanitizeCssColor(raw);
        }
        const scheme = globalXmlParser.element(solidFill, "schemeClr");
        if (scheme) {
            const raw = globalXmlParser.attr(scheme, "val");
            const resolved = resolveSchemeColor(raw);
            return resolved ? sanitizeCssColor(resolved) : null;
        }
        return null;
    }
    function buildCategoryTree(levels, values, colors) {
        const root = {
            label: "", value: 0, color: null, children: [], level: -1, leafIndex: -1,
        };
        if (levels.length === 0)
            return root;
        const perLevel = [];
        const lvl0 = levels[0];
        const level0Nodes = lvl0.points.map((label, idx) => {
            const isLeaf = levels.length === 1;
            const leafIndex = isLeaf ? idx : -1;
            const raw = isLeaf ? values[idx] : NaN;
            const value = Number.isFinite(raw) && raw > 0 ? raw : 0;
            const color = isLeaf ? (colors.get(idx) ?? null) : null;
            return { label, value, color, children: [], level: 0, leafIndex };
        });
        for (const node of level0Nodes)
            root.children.push(node);
        perLevel.push(level0Nodes);
        for (let d = 1; d < levels.length; d++) {
            const lvl = levels[d];
            const isDeepest = d === levels.length - 1;
            const parentLevel = perLevel[d - 1];
            const levelNodes = [];
            for (let idx = 0; idx < lvl.points.length; idx++) {
                const label = lvl.points[idx];
                const parentIdx = lvl.parents[idx];
                const leafIndex = isDeepest ? idx : -1;
                let color = null;
                if (isDeepest) {
                    color = colors.get(idx) ?? null;
                }
                const raw = isDeepest ? values[idx] : NaN;
                const value = Number.isFinite(raw) && raw > 0 ? raw : 0;
                const node = {
                    label, value, color, children: [], level: d, leafIndex,
                };
                const parent = parentIdx >= 0 && parentIdx < parentLevel.length
                    ? parentLevel[parentIdx]
                    : null;
                if (parent) {
                    parent.children.push(node);
                }
                else {
                    root.children.push(node);
                }
                levelNodes.push(node);
            }
            perLevel.push(levelNodes);
        }
        sumValues(root);
        propagateColors(root, null);
        return root;
    }
    function sumValues(node) {
        if (node.children.length === 0) {
            return node.value;
        }
        let total = 0;
        for (const child of node.children) {
            total += sumValues(child);
        }
        if (total > node.value)
            node.value = total;
        return node.value;
    }
    function propagateColors(node, inherited) {
        const effective = node.color ?? inherited;
        if (node.color == null && inherited != null) {
            node.color = inherited;
        }
        for (const child of node.children)
            propagateColors(child, effective);
    }
    function computeMaxDepth(root) {
        let max = 0;
        function visit(n, d) {
            if (d > max)
                max = d;
            for (const c of n.children)
                visit(c, d + 1);
        }
        for (const c of root.children)
            visit(c, 1);
        return max;
    }
    function deriveKey(path) {
        const segs = path.split("/");
        const file = segs[segs.length - 1] ?? "";
        const dot = file.lastIndexOf(".");
        return dot > 0 ? file.slice(0, dot) : file;
    }
    function extractTitle(chart) {
        const titleEl = globalXmlParser.element(chart, "title");
        if (!titleEl)
            return "";
        const parts = [];
        walkTextRuns(titleEl, parts);
        return parts.join("");
    }
    function walkTextRuns(node, out) {
        for (const c of globalXmlParser.elements(node)) {
            if (c.localName === "t" || c.localName === "v") {
                out.push(c.textContent ?? "");
            }
            else {
                walkTextRuns(c, out);
            }
        }
    }

    const SMARTART_LAYOUT_URN_RE = /^urn:microsoft\.com\/office\/officeart\/\d{4}\/\d+\/layout\/[a-z0-9]+$/i;
    class DiagramLayoutPart extends Part {
        constructor(pkg, path) {
            super(pkg, path);
            this.layoutId = "";
        }
        parseXml(root) {
            let uniqueId = globalXmlParser.attr(root, "uniqueId");
            if (!uniqueId) {
                const def = firstDescendant(root, "layoutDef");
                if (def)
                    uniqueId = globalXmlParser.attr(def, "uniqueId");
            }
            if (uniqueId && SMARTART_LAYOUT_URN_RE.test(uniqueId)) {
                this.layoutId = uniqueId;
            }
        }
    }
    class DiagramDataPart extends Part {
        constructor(pkg, path) {
            super(pkg, path);
        }
    }
    class DiagramQuickStylePart extends Part {
        constructor(pkg, path) {
            super(pkg, path);
        }
    }
    class DiagramColorsPart extends Part {
        constructor(pkg, path) {
            super(pkg, path);
        }
    }
    class DiagramDrawingPart extends Part {
        constructor(pkg, path) {
            super(pkg, path);
        }
    }
    function firstDescendant(root, localName) {
        for (const c of globalXmlParser.elements(root)) {
            if (c.localName === localName)
                return c;
            const d = firstDescendant(c, localName);
            if (d)
                return d;
        }
        return null;
    }

    const PLACEHOLDER_WIDTH = 200;
    const PLACEHOLDER_HEIGHT = 100;
    const DECODE_WIDTH = 1000;
    const DECODE_HEIGHT = 800;
    async function convertVectorImage(blob, format) {
        const decoder = resolveDecoder(format);
        if (decoder) {
            try {
                const svg = await decodeToSvg(blob, decoder);
                if (svg)
                    return toObjectURL(svg);
            }
            catch {
            }
        }
        return toObjectURL(placeholderSvg());
    }
    function resolveDecoder(format) {
        const g = globalThis;
        if (format === 'wmf' && g.WMFJS?.Renderer)
            return g.WMFJS;
        if (format === 'emf' && g.EMFJS?.Renderer)
            return g.EMFJS;
        return null;
    }
    async function decodeToSvg(blob, decoder) {
        const buffer = await blob.arrayBuffer();
        const renderer = new decoder.Renderer(new Uint8Array(buffer));
        const result = renderer.render({
            width: `${DECODE_WIDTH}px`,
            height: `${DECODE_HEIGHT}px`,
            xExt: DECODE_WIDTH,
            yExt: DECODE_HEIGHT,
            mapMode: 8,
        });
        const svgEl = result?.tagName?.toLowerCase() === 'svg'
            ? result
            : (result?.firstChild ?? null);
        if (!svgEl || svgEl.tagName?.toLowerCase() !== 'svg') {
            return null;
        }
        svgEl.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
        svgEl.removeAttribute('width');
        svgEl.removeAttribute('height');
        return svgEl.outerHTML ?? null;
    }
    function placeholderSvg() {
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
    function toObjectURL(svg) {
        const prelude = '<?xml version="1.0" encoding="UTF-8"?>';
        const blob = new Blob([prelude, svg], { type: 'image/svg+xml' });
        return URL.createObjectURL(blob);
    }
    function detectVectorFormat(path) {
        if (!path)
            return null;
        const m = path.toLowerCase().match(/\.([a-z0-9]+)$/);
        const ext = m?.[1];
        return ext === 'wmf' || ext === 'emf' ? ext : null;
    }

    class GlossaryDocumentPart extends Part {
        constructor(pkg, path, parser) {
            super(pkg, path);
            this._documentParser = parser;
        }
        parseXml(root) {
            this.body = this._documentParser.parseDocumentFile(root);
        }
    }
    class CustomXmlPart extends Part {
        async load() {
            this.rels = await this._package.loadRelationships(this.path);
            const xmlText = await this._package.load(this.path);
            if (xmlText) {
                try {
                    this.xmlDoc = this._package.parseXmlDocument(xmlText);
                }
                catch {
                    this.xmlDoc = null;
                }
            }
        }
        setItemId(id) {
            this.itemId = id ?? null;
        }
    }
    class CustomXmlPropsPart extends Part {
        parseXml(root) {
            const attrs = Array.from(root.attributes ?? []);
            const idAttr = attrs.find(a => a.localName === "itemID");
            this.itemId = idAttr?.value ?? null;
        }
    }
    const topLevelRels = [
        { type: RelationshipTypes.OfficeDocument, target: "word/document.xml" },
        { type: RelationshipTypes.ExtendedProperties, target: "docProps/app.xml" },
        { type: RelationshipTypes.CoreProperties, target: "docProps/core.xml" },
        { type: RelationshipTypes.CustomProperties, target: "docProps/custom.xml" },
    ];
    [
        { type: RelationshipTypes.GlossaryDocument, target: "glossary/document.xml" },
    ];
    class WordDocument {
        constructor() {
            this.parts = [];
            this.partsMap = {};
            this.contentTypes = [];
            this.customXmlParts = [];
        }
        static async load(blob, parser, options) {
            var d = new WordDocument();
            d._options = options;
            d._parser = parser;
            d._package = await OpenXmlPackage.load(blob, options);
            d.rels = await d._package.loadRelationships();
            d.contentTypes = await d._package.loadContentTypes();
            await Promise.all(topLevelRels.map(rel => {
                const r = d.rels.find(x => x.type === rel.type) ?? rel;
                return d.loadRelationshipPart(r.target, r.type);
            }));
            for (const part of d.customXmlParts) {
                const propsRel = part.rels?.find(r => r.type === RelationshipTypes.CustomXmlProps);
                if (!propsRel)
                    continue;
                const [partFolder] = splitPath(part.path);
                const propsPath = resolvePath(propsRel.target, partFolder);
                const propsPart = d.partsMap[propsPath];
                if (propsPart?.itemId) {
                    part.setItemId(propsPart.itemId);
                }
            }
            if (d.commentsPart) {
                const extComments = d.commentsExtendedPart?.comments ?? [];
                d.commentsPart.buildThreading(extComments);
            }
            return d;
        }
        findCustomXmlByStoreItemId(storeItemID) {
            if (!storeItemID)
                return null;
            const norm = normalizeGuid(storeItemID);
            for (const part of this.customXmlParts) {
                if (!part.xmlDoc)
                    continue;
                if (normalizeGuid(part.itemId) === norm)
                    return part;
            }
            return null;
        }
        get glossaryDocument() {
            return this.glossaryDocumentPart?.body;
        }
        save(type = "blob") {
            return this._package.save(type);
        }
        async loadRelationshipPart(path, type) {
            if (this.partsMap[path])
                return this.partsMap[path];
            if (!this._package.get(path))
                return null;
            let part = null;
            switch (type) {
                case RelationshipTypes.OfficeDocument:
                    this.documentPart = part = new DocumentPart(this._package, path, this._parser);
                    break;
                case RelationshipTypes.FontTable:
                    this.fontTablePart = part = new FontTablePart(this._package, path);
                    break;
                case RelationshipTypes.Numbering:
                    this.numberingPart = part = new NumberingPart(this._package, path, this._parser);
                    break;
                case RelationshipTypes.Styles:
                    this.stylesPart = part = new StylesPart(this._package, path, this._parser);
                    break;
                case RelationshipTypes.Theme:
                    this.themePart = part = new ThemePart(this._package, path);
                    break;
                case RelationshipTypes.Footnotes:
                    this.footnotesPart = part = new FootnotesPart(this._package, path, this._parser);
                    break;
                case RelationshipTypes.Endnotes:
                    this.endnotesPart = part = new EndnotesPart(this._package, path, this._parser);
                    break;
                case RelationshipTypes.Footer:
                    part = new FooterPart(this._package, path, this._parser);
                    break;
                case RelationshipTypes.Header:
                    part = new HeaderPart(this._package, path, this._parser);
                    break;
                case RelationshipTypes.CoreProperties:
                    this.corePropsPart = part = new CorePropsPart(this._package, path);
                    break;
                case RelationshipTypes.ExtendedProperties:
                    this.extendedPropsPart = part = new ExtendedPropsPart(this._package, path);
                    break;
                case RelationshipTypes.CustomProperties:
                    part = new CustomPropsPart(this._package, path);
                    break;
                case RelationshipTypes.Settings:
                    this.settingsPart = part = new SettingsPart(this._package, path);
                    break;
                case RelationshipTypes.Comments:
                    this.commentsPart = part = new CommentsPart(this._package, path, this._parser);
                    break;
                case RelationshipTypes.CommentsExtended:
                    this.commentsExtendedPart = part = new CommentsExtendedPart(this._package, path);
                    break;
                case RelationshipTypes.Chart:
                    part = new ChartPart(this._package, path);
                    break;
                case RelationshipTypes.ChartEx:
                    part = new ChartExPart(this._package, path);
                    break;
                case RelationshipTypes.DiagramLayout:
                    part = new DiagramLayoutPart(this._package, path);
                    break;
                case RelationshipTypes.DiagramData:
                    part = new DiagramDataPart(this._package, path);
                    break;
                case RelationshipTypes.DiagramQuickStyle:
                    part = new DiagramQuickStylePart(this._package, path);
                    break;
                case RelationshipTypes.DiagramColors:
                    part = new DiagramColorsPart(this._package, path);
                    break;
                case RelationshipTypes.DiagramDrawing:
                    part = new DiagramDrawingPart(this._package, path);
                    break;
                case RelationshipTypes.GlossaryDocument:
                    this.glossaryDocumentPart = part = new GlossaryDocumentPart(this._package, path, this._parser);
                    break;
                case RelationshipTypes.CustomXml: {
                    const xmlPart = new CustomXmlPart(this._package, path);
                    this.customXmlParts.push(xmlPart);
                    part = xmlPart;
                    break;
                }
                case RelationshipTypes.CustomXmlProps:
                    part = new CustomXmlPropsPart(this._package, path);
                    break;
            }
            if (part == null)
                return Promise.resolve(null);
            this.partsMap[path] = part;
            this.parts.push(part);
            await part.load();
            if (part.rels?.length > 0) {
                const [folder] = splitPath(part.path);
                await Promise.all(part.rels.map(rel => this.loadRelationshipPart(resolvePath(rel.target, folder), rel.type)));
            }
            return part;
        }
        async loadDocumentImage(id, part) {
            const path = this.getPathById(part ?? this.documentPart, id);
            if (!path)
                return null;
            const blob = await this._package.load(path, "blob");
            return this.blobToImageURL(blob, path);
        }
        async loadNumberingImage(id) {
            const path = this.getPathById(this.numberingPart, id);
            if (!path)
                return null;
            const blob = await this._package.load(path, "blob");
            return this.blobToImageURL(blob, path);
        }
        async blobToImageURL(blob, path) {
            if (!blob)
                return null;
            const vector = detectVectorFormat(path);
            if (vector) {
                return convertVectorImage(blob, vector);
            }
            const url = this.blobToURL(blob, path);
            return typeof url === 'string' ? url : await url;
        }
        async loadFont(id, key) {
            const path = this.getPathById(this.fontTablePart, id);
            if (!path)
                return null;
            const x = await this._package.load(path, "uint8array");
            return x ? this.blobToURL(new Blob([deobfuscate(x, key)]), path) : x;
        }
        async loadAltChunk(id, part) {
            const path = this.getPathById(part ?? this.documentPart, id);
            return path ? this._package.load(path, "string") : Promise.resolve(null);
        }
        blobToURL(blob, path) {
            if (!blob)
                return null;
            if (path) {
                const ct = this.contentTypes.find(x => x.partName === path || (x.extension && path.endsWith(`.${x.extension}`)));
                blob = ct ? new Blob([blob], { type: ct.contentType }) : blob;
            }
            if (this._options.useBase64URL) {
                return blobToBase64(blob);
            }
            return URL.createObjectURL(blob);
        }
        findPartByRelId(id, basePart = null) {
            var rel = (basePart.rels ?? this.rels).find(r => r.id == id);
            const folder = basePart ? splitPath(basePart.path)[0] : '';
            return rel ? this.partsMap[resolvePath(rel.target, folder)] : null;
        }
        getPathById(part, id) {
            const rel = part.rels.find(x => x.id == id);
            const [folder] = splitPath(part.path);
            return rel ? resolvePath(rel.target, folder) : null;
        }
    }
    function normalizeGuid(s) {
        if (!s)
            return null;
        return s.replace(/[{}]/g, '').toUpperCase();
    }
    function deobfuscate(data, guidKey) {
        const len = 16;
        const trimmed = guidKey.replace(/{|}|-/g, "");
        const numbers = new Array(len);
        for (let i = 0; i < len; i++)
            numbers[len - i - 1] = parseInt(trimmed.substring(i * 2, i * 2 + 2), 16);
        for (let i = 0; i < 32; i++)
            data[i] = data[i] ^ numbers[i % len];
        return data;
    }

    function parseBookmarkStart(elem, xml) {
        return {
            type: DomType.BookmarkStart,
            id: xml.attr(elem, "id"),
            name: xml.attr(elem, "name"),
            colFirst: xml.intAttr(elem, "colFirst"),
            colLast: xml.intAttr(elem, "colLast")
        };
    }
    function parseBookmarkEnd(elem, xml) {
        return {
            type: DomType.BookmarkEnd,
            id: xml.attr(elem, "id")
        };
    }

    function sanitizeVmlColor(value) {
        if (typeof value !== 'string')
            return null;
        const stripped = value.replace(/\s*\[\d+\]\s*$/, '');
        return sanitizeCssColor(stripped);
    }
    const SAFE_VML_PATH = /^[0-9eEmMlLcCxX.,\-\s]*$/;
    let vmlDefsCounter = 0;
    function nextVmlId(prefix) {
        return `${prefix}-${++vmlDefsCounter}`;
    }
    class VmlElement extends OpenXmlElementBase {
        constructor() {
            super(...arguments);
            this.type = DomType.VmlElement;
            this.attrs = {};
        }
    }
    function makeVml(tagName, attrs = {}, children = []) {
        const v = new VmlElement();
        v.tagName = tagName;
        v.attrs = attrs;
        for (const c of children)
            v.children.push(c);
        return v;
    }
    function parseVmlElement(elem, parser) {
        var result = new VmlElement();
        switch (elem.localName) {
            case "rect":
                result.tagName = "rect";
                Object.assign(result.attrs, { width: '100%', height: '100%' });
                break;
            case "oval":
                result.tagName = "ellipse";
                Object.assign(result.attrs, { cx: "50%", cy: "50%", rx: "50%", ry: "50%" });
                break;
            case "line":
                result.tagName = "line";
                break;
            case "shape":
                result.tagName = "g";
                break;
            case "textbox":
                result.tagName = "foreignObject";
                Object.assign(result.attrs, { width: '100%', height: '100%' });
                break;
            case "group":
                result.tagName = "svg";
                applyGroupCoordSystem(elem, result);
                break;
            case "path": {
                result.tagName = "path";
                const rawPath = globalXmlParser.attr(elem, "v");
                const safePath = convertVmlPathToSvg(rawPath);
                if (safePath)
                    result.attrs.d = safePath;
                break;
            }
            case "extrusion":
                return null;
            default:
                return null;
        }
        for (const at of globalXmlParser.attrs(elem)) {
            switch (at.localName) {
                case "style":
                    result.cssStyleText = at.value;
                    break;
                case "fillcolor": {
                    const fill = sanitizeVmlColor(at.value);
                    if (fill)
                        result.attrs.fill = fill;
                    break;
                }
                case "from":
                    const [x1, y1] = parsePoint(at.value);
                    Object.assign(result.attrs, { x1, y1 });
                    break;
                case "to":
                    const [x2, y2] = parsePoint(at.value);
                    Object.assign(result.attrs, { x2, y2 });
                    break;
            }
        }
        const defsChildren = [];
        for (const el of globalXmlParser.elements(elem)) {
            switch (el.localName) {
                case "stroke":
                    Object.assign(result.attrs, parseStroke(el));
                    break;
                case "fill": {
                    const { attrs, defs } = parseFill(el);
                    Object.assign(result.attrs, attrs);
                    if (defs)
                        defsChildren.push(defs);
                    break;
                }
                case "shadow": {
                    const shadow = parseShadow(el);
                    if (shadow) {
                        defsChildren.push(shadow.defs);
                        if (!result.attrs.filter) {
                            result.attrs.filter = `url(#${shadow.id})`;
                        }
                    }
                    break;
                }
                case "imagedata":
                    result.tagName = "image";
                    Object.assign(result.attrs, { width: '100%', height: '100%' });
                    result.imageHref = {
                        id: globalXmlParser.attr(el, "id"),
                        title: globalXmlParser.attr(el, "title"),
                    };
                    break;
                case "txbxContent":
                    result.children.push(...parser.parseBodyElements(el));
                    break;
                default:
                    const child = parseVmlElement(el, parser);
                    child && result.children.push(child);
                    break;
            }
        }
        if (elem.localName === "group") {
            rewriteGroupChildPositions(elem, result);
        }
        if (defsChildren.length) {
            const defs = makeVml("defs", {}, defsChildren);
            result.children.unshift(defs);
        }
        return result;
    }
    function parseStroke(el) {
        const result = {
            'stroke-width': globalXmlParser.lengthAttr(el, "weight", LengthUsage.Emu) ?? '1px'
        };
        const stroke = sanitizeVmlColor(globalXmlParser.attr(el, "color"));
        if (stroke)
            result['stroke'] = stroke;
        return result;
    }
    function parseFill(el) {
        const type = globalXmlParser.attr(el, "type");
        const attrs = {};
        if (!type || type === "solid") {
            const color = sanitizeVmlColor(globalXmlParser.attr(el, "color"));
            if (color)
                attrs.fill = color;
            return { attrs };
        }
        if (type === "gradient" || type === "gradientRadial") {
            return parseGradientFill(el, type);
        }
        if (type === "pattern" || type === "tile") {
            return parsePatternFill(el);
        }
        return { attrs };
    }
    function parseGradientFill(el, type) {
        const color1 = sanitizeVmlColor(globalXmlParser.attr(el, "color")) ?? "#000000";
        const color2 = sanitizeVmlColor(globalXmlParser.attr(el, "color2")) ?? "#FFFFFF";
        const rawAngle = parseFloat(globalXmlParser.attr(el, "angle"));
        const angle = Number.isFinite(rawAngle) ? rawAngle : 0;
        const rawFocus = parseFloat(globalXmlParser.attr(el, "focus"));
        const focus = Number.isFinite(rawFocus) ? clampNum(rawFocus, -100, 100) : 0;
        const id = nextVmlId("vml-grad");
        const rad = (angle - 90) * Math.PI / 180;
        const cx = 0.5, cy = 0.5;
        const x1 = cx - Math.cos(rad) * 0.5;
        const y1 = cy - Math.sin(rad) * 0.5;
        const x2 = cx + Math.cos(rad) * 0.5;
        const y2 = cy + Math.sin(rad) * 0.5;
        const stops = [];
        if (focus === 0) {
            stops.push(makeVml("stop", { offset: "0%", "stop-color": color1 }));
            stops.push(makeVml("stop", { offset: "100%", "stop-color": color2 }));
        }
        else {
            const mid = `${(50 + focus / 2).toFixed(2)}%`;
            stops.push(makeVml("stop", { offset: "0%", "stop-color": color1 }));
            stops.push(makeVml("stop", { offset: mid, "stop-color": color2 }));
            stops.push(makeVml("stop", { offset: "100%", "stop-color": color1 }));
        }
        const gradientTag = type === "gradientRadial" ? "radialGradient" : "linearGradient";
        const gradAttrs = { id };
        if (gradientTag === "linearGradient") {
            Object.assign(gradAttrs, {
                x1: x1.toFixed(4),
                y1: y1.toFixed(4),
                x2: x2.toFixed(4),
                y2: y2.toFixed(4),
            });
        }
        const defs = makeVml(gradientTag, gradAttrs, stops);
        return {
            attrs: { fill: `url(#${id})` },
            defs,
        };
    }
    function parsePatternFill(el) {
        const color = sanitizeVmlColor(globalXmlParser.attr(el, "color"));
        const color2 = sanitizeVmlColor(globalXmlParser.attr(el, "color2"));
        const id = nextVmlId("vml-pat");
        if (!color && !color2) {
            return { attrs: {} };
        }
        const size = 8;
        const bg = makeVml("rect", {
            x: "0", y: "0",
            width: `${size}`, height: `${size}`,
            fill: color2 ?? "#FFFFFF",
        });
        const stripe = makeVml("path", {
            d: `M0,${size} L${size},0`,
            stroke: color ?? "#000000",
            "stroke-width": "1",
        });
        const defs = makeVml("pattern", {
            id,
            x: "0", y: "0",
            width: `${size}`, height: `${size}`,
            patternUnits: "userSpaceOnUse",
        }, [bg, stripe]);
        return {
            attrs: { fill: `url(#${id})` },
            defs,
        };
    }
    function parseShadow(el) {
        const on = globalXmlParser.attr(el, "on");
        if (on && !/^(t|true|1|on)$/i.test(on))
            return null;
        const color = sanitizeVmlColor(globalXmlParser.attr(el, "color")) ?? "#000000";
        const opacityRaw = globalXmlParser.attr(el, "opacity");
        const opacity = parseVmlOpacity(opacityRaw);
        const [dx, dy] = parseVmlOffset(globalXmlParser.attr(el, "offset"));
        const id = nextVmlId("vml-shadow");
        const feAttrs = {
            dx: dx.toFixed(2),
            dy: dy.toFixed(2),
            stdDeviation: "0",
            "flood-color": color,
            "flood-opacity": opacity.toFixed(3),
        };
        const fe = makeVml("feDropShadow", feAttrs);
        const filter = makeVml("filter", {
            id,
            x: "-50%",
            y: "-50%",
            width: "200%",
            height: "200%",
        }, [fe]);
        return { id, defs: filter };
    }
    function parseVmlOpacity(val) {
        if (!val)
            return 1;
        const s = val.trim();
        if (/f$/i.test(s)) {
            const n = parseFloat(s);
            return Number.isFinite(n) ? clampNum(n / 65536, 0, 1) : 1;
        }
        if (s.endsWith('%')) {
            const n = parseFloat(s);
            return Number.isFinite(n) ? clampNum(n / 100, 0, 1) : 1;
        }
        const n = parseFloat(s);
        return Number.isFinite(n) ? clampNum(n, 0, 1) : 1;
    }
    function parseVmlOffset(val) {
        if (!val)
            return [2, 2];
        const parts = val.split(',').map(p => p.trim());
        const dx = parseVmlLengthToPt(parts[0]);
        const dy = parseVmlLengthToPt(parts[1] ?? parts[0]);
        return [dx, dy];
    }
    function parseVmlLengthToPt(s) {
        if (!s)
            return 0;
        const n = parseFloat(s);
        if (!Number.isFinite(n))
            return 0;
        if (/pt$/i.test(s))
            return n;
        if (/px$/i.test(s))
            return n * 0.75;
        if (/in$/i.test(s))
            return n * 72;
        if (/cm$/i.test(s))
            return n * 28.3464567;
        if (/mm$/i.test(s))
            return n * 2.8346457;
        return n;
    }
    function parsePoint(val) {
        return val.split(",");
    }
    function convertVmlPathToSvg(path) {
        if (!path)
            return null;
        if (!SAFE_VML_PATH.test(path))
            return null;
        const cmdMap = {
            m: 'M', M: 'M',
            l: 'L', L: 'L',
            c: 'C', C: 'C',
            x: 'Z', X: 'Z',
            e: '', E: '',
        };
        const out = [];
        const re = /([mMlLcCxXeE])|(-?\d+(?:\.\d+)?)|([,\s])/g;
        let match;
        while ((match = re.exec(path)) !== null) {
            if (match[1] !== undefined) {
                const c = cmdMap[match[1]];
                if (c)
                    out.push(c);
            }
            else if (match[2] !== undefined) {
                out.push(convertLength(match[2], LengthUsage.VmlEmu));
            }
            else if (match[3] !== undefined) {
                if (out.length && !/[,\s]$/.test(out[out.length - 1])) {
                    out.push(' ');
                }
            }
        }
        const joined = out.join('');
        return joined.replace(/\s+/g, ' ').replace(/\s*,\s*/g, ',').trim() || null;
    }
    function applyGroupCoordSystem(elem, result) {
        const [csx, csy] = parseCoordPair(globalXmlParser.attr(elem, "coordsize")) ?? [1000, 1000];
        const [cox, coy] = parseCoordPair(globalXmlParser.attr(elem, "coordorigin")) ?? [0, 0];
        result.attrs.viewBox = `${cox} ${coy} ${csx} ${csy}`;
        result.attrs.preserveAspectRatio = "none";
        result.__groupCoord = { csx, csy, cox, coy };
    }
    function rewriteGroupChildPositions(_groupElem, group) {
        for (const child of group.children) {
            if (!(child instanceof VmlElement))
                continue;
            const style = child.cssStyleText;
            if (!style)
                continue;
            const rules = parseCssRules(style);
            const left = parsePositionValue(rules.left);
            const top = parsePositionValue(rules.top);
            const width = parsePositionValue(rules.width);
            const height = parsePositionValue(rules.height);
            switch (child.tagName) {
                case "rect":
                case "image":
                case "foreignObject":
                case "svg":
                    if (left != null)
                        child.attrs.x = left.toString();
                    if (top != null)
                        child.attrs.y = top.toString();
                    if (width != null)
                        child.attrs.width = width.toString();
                    if (height != null)
                        child.attrs.height = height.toString();
                    break;
                case "ellipse":
                    if (left != null && width != null) {
                        child.attrs.cx = (left + width / 2).toString();
                        child.attrs.rx = (width / 2).toString();
                    }
                    if (top != null && height != null) {
                        child.attrs.cy = (top + height / 2).toString();
                        child.attrs.ry = (height / 2).toString();
                    }
                    break;
                case "g":
                    if (left != null || top != null) {
                        const tx = left ?? 0;
                        const ty = top ?? 0;
                        const existing = child.attrs.transform ?? '';
                        child.attrs.transform = `translate(${tx} ${ty}) ${existing}`.trim();
                    }
                    break;
            }
        }
    }
    function parseCoordPair(val) {
        if (!val)
            return null;
        const parts = val.split(',').map(s => parseFloat(s.trim()));
        if (parts.length !== 2 || !Number.isFinite(parts[0]) || !Number.isFinite(parts[1]))
            return null;
        return [parts[0], parts[1]];
    }
    function parsePositionValue(val) {
        if (val == null)
            return null;
        const n = parseFloat(val);
        return Number.isFinite(n) ? n : null;
    }
    function clampNum(val, min, max) {
        return val < min ? min : val > max ? max : val;
    }

    class WmlComment extends OpenXmlElementBase {
        constructor() {
            super(...arguments);
            this.type = DomType.Comment;
            this.done = false;
            this.parentCommentId = null;
            this.replies = [];
        }
    }
    class WmlCommentReference extends OpenXmlElementBase {
        constructor(id) {
            super();
            this.id = id;
            this.type = DomType.CommentReference;
        }
    }
    class WmlCommentRangeStart extends OpenXmlElementBase {
        constructor(id) {
            super();
            this.id = id;
            this.type = DomType.CommentRangeStart;
        }
    }
    class WmlCommentRangeEnd extends OpenXmlElementBase {
        constructor(id) {
            super();
            this.id = id;
            this.type = DomType.CommentRangeEnd;
        }
    }

    function parseRevisionAttrs(elem) {
        return {
            id: globalXmlParser.attr(elem, "id"),
            author: globalXmlParser.attr(elem, "author"),
            date: globalXmlParser.attr(elem, "date")
        };
    }
    const FORMATTING_PROP_NAMES = {
        b: "bold", i: "italic", u: "underline", strike: "strikethrough",
        sz: "font size", rFonts: "font", color: "color", highlight: "highlight",
        jc: "alignment", ind: "indent", spacing: "spacing", numPr: "numbering",
        pStyle: "style", rStyle: "style"
    };
    function parseFormattingRevision(elem) {
        const rev = {
            id: globalXmlParser.attr(elem, "id"),
            author: globalXmlParser.attr(elem, "author"),
            date: globalXmlParser.attr(elem, "date"),
            changedProps: []
        };
        const prev = globalXmlParser.elements(elem).find(e => e.localName === "rPr" || e.localName === "pPr");
        if (prev) {
            const seen = new Set();
            for (const child of globalXmlParser.elements(prev)) {
                const pretty = FORMATTING_PROP_NAMES[child.localName] ?? child.localName;
                if (seen.has(pretty))
                    continue;
                seen.add(pretty);
                rev.changedProps.push(pretty);
                if (rev.changedProps.length >= 5)
                    break;
            }
        }
        return rev;
    }
    function classNameOfCnfStyle(c) {
        const val = globalXmlParser.attr(c, "val");
        if (!val)
            return '';
        const classes = [
            'first-row', 'last-row', 'first-col', 'last-col',
            'odd-col', 'even-col', 'odd-row', 'even-row',
            'ne-cell', 'nw-cell', 'se-cell', 'sw-cell'
        ];
        return classes.filter((_, i) => val[i] == '1').join(' ');
    }
    var autos = {
        shd: "inherit",
        color: "black",
        borderColor: "black",
        highlight: "transparent"
    };
    const supportedNamespaceURIs = [
        "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
        "http://schemas.microsoft.com/office/word/2010/wordml",
    ];
    function findAncestorByLocalName(start, localName) {
        let n = start ? start.parentNode : null;
        while (n && n.nodeType === 1) {
            if (n.localName === localName)
                return n;
            n = n.parentNode;
        }
        return null;
    }
    const mmlTagMap = {
        "oMath": DomType.MmlMath,
        "oMathPara": DomType.MmlMathParagraph,
        "f": DomType.MmlFraction,
        "func": DomType.MmlFunction,
        "fName": DomType.MmlFunctionName,
        "num": DomType.MmlNumerator,
        "den": DomType.MmlDenominator,
        "rad": DomType.MmlRadical,
        "deg": DomType.MmlDegree,
        "e": DomType.MmlBase,
        "sSup": DomType.MmlSuperscript,
        "sSub": DomType.MmlSubscript,
        "sPre": DomType.MmlPreSubSuper,
        "sup": DomType.MmlSuperArgument,
        "sub": DomType.MmlSubArgument,
        "d": DomType.MmlDelimiter,
        "nary": DomType.MmlNary,
        "eqArr": DomType.MmlEquationArray,
        "lim": DomType.MmlLimit,
        "limLow": DomType.MmlLimitLower,
        "m": DomType.MmlMatrix,
        "mr": DomType.MmlMatrixRow,
        "box": DomType.MmlBox,
        "bar": DomType.MmlBar,
        "groupChr": DomType.MmlGroupChar,
        "acc": DomType.MmlAccent,
        "borderBox": DomType.MmlBorderBox,
        "sSubSup": DomType.MmlSubSuperscript,
        "phant": DomType.MmlPhantom,
        "sGroup": DomType.MmlGroup
    };
    class DocumentParser {
        constructor(options) {
            this.PRESET_GEOMETRY_ALLOWLIST = new Set([
                "rect", "roundRect", "ellipse", "triangle", "rtTriangle", "diamond",
                "parallelogram", "trapezoid", "pentagon", "hexagon", "octagon", "line",
                "rightArrow", "leftArrow", "upArrow", "downArrow", "leftRightArrow",
                "wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout",
                "star5", "star6", "star8", "cloudCallout",
            ]);
            this.PATTERN_ALLOWLIST = new Set([
                "dkDnDiag", "ltDnDiag", "dkUpDiag", "ltUpDiag",
                "dkHorz", "ltHorz", "dkVert", "ltVert",
                "cross", "diagCross",
            ]);
            this.ADJUSTMENT_NAME_ALLOWLIST = new Set([
                "adj", "adj1", "adj2", "adj3", "adj4", "adj5", "adj6", "adj7", "adj8",
            ]);
            this.options = {
                ignoreWidth: false,
                debug: false,
                ...options
            };
        }
        parseNotes(xmlDoc, elemName, elemClass) {
            var result = [];
            for (let el of globalXmlParser.elements(xmlDoc, elemName)) {
                const node = new elemClass();
                node.id = globalXmlParser.attr(el, "id");
                node.noteType = globalXmlParser.attr(el, "type");
                node.children = this.parseBodyElements(el);
                result.push(node);
            }
            return result;
        }
        parseComments(xmlDoc) {
            var result = [];
            for (let el of globalXmlParser.elements(xmlDoc, "comment")) {
                const item = new WmlComment();
                item.id = globalXmlParser.attr(el, "id");
                item.author = globalXmlParser.attr(el, "author");
                item.initials = globalXmlParser.attr(el, "initials");
                item.date = globalXmlParser.attr(el, "date");
                const paraId = el.getAttributeNS("http://schemas.microsoft.com/office/word/2010/wordml", "paraId")
                    ?? el.getAttribute("w14:paraId")
                    ?? globalXmlParser.attr(el, "paraId");
                if (paraId) {
                    item.paraId = paraId;
                }
                item.children = this.parseBodyElements(el);
                result.push(item);
            }
            return result;
        }
        parseDocumentFile(xmlDoc) {
            var xbody = globalXmlParser.element(xmlDoc, "body");
            var background = globalXmlParser.element(xmlDoc, "background");
            var sectPr = globalXmlParser.element(xbody, "sectPr");
            return {
                type: DomType.Document,
                children: this.parseBodyElements(xbody),
                props: sectPr ? parseSectionProperties(sectPr, globalXmlParser) : {},
                cssStyle: background ? this.parseBackground(background) : {},
            };
        }
        parseBackground(elem) {
            var result = {};
            var color = xmlUtil.colorAttr(elem, "color");
            if (color) {
                result["background-color"] = color;
            }
            return result;
        }
        parseBodyElements(element) {
            var children = [];
            for (const elem of globalXmlParser.elements(element)) {
                switch (elem.localName) {
                    case "p":
                        children.push(this.parseParagraph(elem));
                        break;
                    case "altChunk":
                        children.push(this.parseAltChunk(elem));
                        break;
                    case "tbl":
                        children.push(this.parseTable(elem));
                        break;
                    case "sdt":
                        children.push(...this.parseSdt(elem, e => this.parseBodyElements(e)));
                        break;
                }
            }
            return children;
        }
        parseStylesFile(xstyles) {
            var result = [];
            for (const n of globalXmlParser.elements(xstyles)) {
                switch (n.localName) {
                    case "style":
                        result.push(this.parseStyle(n));
                        break;
                    case "docDefaults":
                        result.push(this.parseDefaultStyles(n));
                        break;
                }
            }
            return result;
        }
        parseDefaultStyles(node) {
            var result = {
                id: null,
                name: null,
                target: null,
                basedOn: null,
                styles: []
            };
            for (const c of globalXmlParser.elements(node)) {
                switch (c.localName) {
                    case "rPrDefault":
                        var rPr = globalXmlParser.element(c, "rPr");
                        if (rPr)
                            result.styles.push({
                                target: "span",
                                values: this.parseDefaultProperties(rPr, {})
                            });
                        break;
                    case "pPrDefault":
                        var pPr = globalXmlParser.element(c, "pPr");
                        if (pPr)
                            result.styles.push({
                                target: "p",
                                values: this.parseDefaultProperties(pPr, {})
                            });
                        break;
                }
            }
            return result;
        }
        parseStyle(node) {
            var result = {
                id: globalXmlParser.attr(node, "styleId"),
                isDefault: globalXmlParser.boolAttr(node, "default"),
                name: null,
                target: null,
                basedOn: null,
                styles: [],
                linked: null
            };
            switch (globalXmlParser.attr(node, "type")) {
                case "paragraph":
                    result.target = "p";
                    break;
                case "table":
                    result.target = "table";
                    break;
                case "character":
                    result.target = "span";
                    break;
            }
            for (const n of globalXmlParser.elements(node)) {
                switch (n.localName) {
                    case "basedOn":
                        result.basedOn = globalXmlParser.attr(n, "val");
                        break;
                    case "name":
                        result.name = globalXmlParser.attr(n, "val");
                        break;
                    case "link":
                        result.linked = globalXmlParser.attr(n, "val");
                        break;
                    case "next":
                        result.next = globalXmlParser.attr(n, "val");
                        break;
                    case "aliases":
                        result.aliases = globalXmlParser.attr(n, "val").split(",");
                        break;
                    case "pPr":
                        result.styles.push({
                            target: "p",
                            values: this.parseDefaultProperties(n, {})
                        });
                        result.paragraphProps = parseParagraphProperties(n, globalXmlParser);
                        break;
                    case "rPr":
                        result.styles.push({
                            target: "span",
                            values: this.parseDefaultProperties(n, {})
                        });
                        result.runProps = parseRunProperties(n, globalXmlParser);
                        break;
                    case "tblPr":
                    case "tcPr":
                        result.styles.push({
                            target: "td",
                            values: this.parseDefaultProperties(n, {})
                        });
                        break;
                    case "tblStylePr":
                        for (let s of this.parseTableStyle(n))
                            result.styles.push(s);
                        break;
                    case "rsid":
                    case "qFormat":
                    case "hidden":
                    case "semiHidden":
                    case "unhideWhenUsed":
                    case "autoRedefine":
                    case "uiPriority":
                        break;
                    default:
                        this.options.debug && console.warn(`DOCX: Unknown style element: ${n.localName}`);
                }
            }
            return result;
        }
        parseTableStyle(node) {
            var result = [];
            var type = globalXmlParser.attr(node, "type");
            var selector = "";
            var modificator = "";
            switch (type) {
                case "firstRow":
                    modificator = ".first-row";
                    selector = "tr.first-row td";
                    break;
                case "lastRow":
                    modificator = ".last-row";
                    selector = "tr.last-row td";
                    break;
                case "firstCol":
                    modificator = ".first-col";
                    selector = "td.first-col";
                    break;
                case "lastCol":
                    modificator = ".last-col";
                    selector = "td.last-col";
                    break;
                case "band1Vert":
                    modificator = ":not(.no-vband)";
                    selector = "td.odd-col";
                    break;
                case "band2Vert":
                    modificator = ":not(.no-vband)";
                    selector = "td.even-col";
                    break;
                case "band1Horz":
                    modificator = ":not(.no-hband)";
                    selector = "tr.odd-row";
                    break;
                case "band2Horz":
                    modificator = ":not(.no-hband)";
                    selector = "tr.even-row";
                    break;
                default: return [];
            }
            for (const n of globalXmlParser.elements(node)) {
                switch (n.localName) {
                    case "pPr":
                        result.push({
                            target: `${selector} p`,
                            mod: modificator,
                            values: this.parseDefaultProperties(n, {})
                        });
                        break;
                    case "rPr":
                        result.push({
                            target: `${selector} span`,
                            mod: modificator,
                            values: this.parseDefaultProperties(n, {})
                        });
                        break;
                    case "tblPr":
                    case "tcPr":
                        result.push({
                            target: selector,
                            mod: modificator,
                            values: this.parseDefaultProperties(n, {})
                        });
                        break;
                }
            }
            return result;
        }
        parseNumberingFile(node) {
            var abstractLevels = {};
            var bullets = [];
            var numElements = [];
            for (const n of globalXmlParser.elements(node)) {
                switch (n.localName) {
                    case "abstractNum":
                        var absId = globalXmlParser.attr(n, "abstractNumId");
                        abstractLevels[absId] = this.parseAbstractNumbering(n, bullets);
                        break;
                    case "numPicBullet":
                        bullets.push(this.parseNumberingPicBullet(n));
                        break;
                    case "num":
                        numElements.push(n);
                        break;
                }
            }
            var result = [];
            for (const n of numElements) {
                var numId = globalXmlParser.attr(n, "numId");
                var abstractNumId = globalXmlParser.elementAttr(n, "abstractNumId", "val");
                var baseLevels = abstractLevels[abstractNumId];
                if (!baseLevels)
                    continue;
                var overrides = {};
                for (const child of globalXmlParser.elements(n, "lvlOverride")) {
                    var ilvl = globalXmlParser.intAttr(child, "ilvl");
                    var entry = {};
                    for (const sub of globalXmlParser.elements(child)) {
                        if (sub.localName === "startOverride") {
                            entry.start = globalXmlParser.intAttr(sub, "val");
                        }
                        else if (sub.localName === "lvl") {
                            entry.level = this.parseNumberingLevel(abstractNumId, sub, bullets);
                        }
                    }
                    overrides[ilvl] = entry;
                }
                for (const base of baseLevels) {
                    var clone = {
                        ...base,
                        pStyle: { ...base.pStyle },
                        rStyle: { ...base.rStyle },
                        id: numId,
                    };
                    var ov = overrides[base.level];
                    if (ov) {
                        if (ov.level) {
                            if (ov.level.start !== undefined)
                                clone.start = ov.level.start;
                            if (ov.level.levelText !== undefined)
                                clone.levelText = ov.level.levelText;
                            if (ov.level.format !== undefined)
                                clone.format = ov.level.format;
                            if (ov.level.suff !== undefined)
                                clone.suff = ov.level.suff;
                            if (ov.level.restart !== undefined)
                                clone.restart = ov.level.restart;
                            if (ov.level.justification !== undefined)
                                clone.justification = ov.level.justification;
                            if (ov.level.isLgl !== undefined)
                                clone.isLgl = ov.level.isLgl;
                            if (ov.level.bullet !== undefined)
                                clone.bullet = ov.level.bullet;
                            if (ov.level.pStyleName !== undefined)
                                clone.pStyleName = ov.level.pStyleName;
                            if (ov.level.pStyle && Object.keys(ov.level.pStyle).length) {
                                clone.pStyle = { ...clone.pStyle, ...ov.level.pStyle };
                            }
                            if (ov.level.rStyle && Object.keys(ov.level.rStyle).length) {
                                clone.rStyle = { ...clone.rStyle, ...ov.level.rStyle };
                            }
                        }
                        if (ov.start !== undefined)
                            clone.start = ov.start;
                    }
                    result.push(clone);
                }
            }
            return result;
        }
        parseNumberingPicBullet(elem) {
            var pict = globalXmlParser.element(elem, "pict");
            var shape = pict && globalXmlParser.element(pict, "shape");
            var imagedata = shape && globalXmlParser.element(shape, "imagedata");
            return imagedata ? {
                id: globalXmlParser.intAttr(elem, "numPicBulletId"),
                src: globalXmlParser.attr(imagedata, "id"),
                style: globalXmlParser.attr(shape, "style")
            } : null;
        }
        parseAbstractNumbering(node, bullets) {
            var result = [];
            var id = globalXmlParser.attr(node, "abstractNumId");
            for (const n of globalXmlParser.elements(node)) {
                switch (n.localName) {
                    case "lvl":
                        result.push(this.parseNumberingLevel(id, n, bullets));
                        break;
                }
            }
            return result;
        }
        parseNumberingLevel(id, node, bullets) {
            var result = {
                id: id,
                level: globalXmlParser.intAttr(node, "ilvl"),
                start: 1,
                pStyleName: undefined,
                pStyle: {},
                rStyle: {},
                suff: "tab"
            };
            for (const n of globalXmlParser.elements(node)) {
                switch (n.localName) {
                    case "start":
                        result.start = globalXmlParser.intAttr(n, "val");
                        break;
                    case "pPr":
                        this.parseDefaultProperties(n, result.pStyle);
                        break;
                    case "rPr":
                        this.parseDefaultProperties(n, result.rStyle);
                        break;
                    case "lvlPicBulletId":
                        var bulletId = globalXmlParser.intAttr(n, "val");
                        result.bullet = bullets.find(x => x?.id == bulletId);
                        break;
                    case "lvlText":
                        result.levelText = globalXmlParser.attr(n, "val");
                        break;
                    case "pStyle":
                        result.pStyleName = globalXmlParser.attr(n, "val");
                        break;
                    case "numFmt":
                        result.format = globalXmlParser.attr(n, "val");
                        break;
                    case "suff":
                        result.suff = globalXmlParser.attr(n, "val");
                        break;
                    case "lvlRestart":
                        result.restart = globalXmlParser.intAttr(n, "val");
                        break;
                    case "lvlJc":
                        result.justification = globalXmlParser.attr(n, "val");
                        break;
                    case "isLgl":
                        var lglVal = globalXmlParser.attr(n, "val");
                        result.isLgl = lglVal === undefined || lglVal === "" ||
                            lglVal === "1" || lglVal === "true" || lglVal === "on";
                        break;
                }
            }
            return result;
        }
        parseSdt(node, parser) {
            const sdtContent = globalXmlParser.element(node, "sdtContent");
            if (!sdtContent)
                return [];
            const children = parser(sdtContent) ?? [];
            const sdtPr = globalXmlParser.element(node, "sdtPr");
            if (sdtPr) {
                const aliasEl = globalXmlParser.element(sdtPr, "alias");
                const tagEl = globalXmlParser.element(sdtPr, "tag");
                const alias = aliasEl ? globalXmlParser.attr(aliasEl, "val") : null;
                const tag = tagEl ? globalXmlParser.attr(tagEl, "val") : null;
                const control = this.parseSdtControl(sdtPr);
                const dataBindingEl = globalXmlParser.element(sdtPr, "dataBinding");
                let dataBinding = null;
                if (dataBindingEl) {
                    const xpath = globalXmlParser.attr(dataBindingEl, "xpath");
                    if (xpath) {
                        dataBinding = { xpath };
                        const storeItemID = globalXmlParser.attr(dataBindingEl, "storeItemID");
                        if (storeItemID)
                            dataBinding.storeItemID = storeItemID;
                        const prefixMappings = globalXmlParser.attr(dataBindingEl, "prefixMappings");
                        if (prefixMappings)
                            dataBinding.prefixMappings = prefixMappings;
                    }
                }
                if (alias || tag || control || dataBinding) {
                    const wrapper = { type: DomType.Sdt, children };
                    if (alias)
                        wrapper.sdtAlias = alias;
                    if (tag)
                        wrapper.sdtTag = tag;
                    if (control)
                        wrapper.sdtControl = control;
                    if (dataBinding)
                        wrapper.dataBinding = dataBinding;
                    return [wrapper];
                }
            }
            return children;
        }
        parseFfData(ffData) {
            const result = {};
            const textInput = globalXmlParser.element(ffData, "textInput");
            if (textInput) {
                result.formFieldType = 'text';
                const defaultEl = globalXmlParser.element(textInput, "default");
                if (defaultEl) {
                    const v = globalXmlParser.attr(defaultEl, "val");
                    if (v != null)
                        result.defaultText = v;
                }
                const maxLengthEl = globalXmlParser.element(textInput, "maxLength");
                if (maxLengthEl) {
                    const v = globalXmlParser.intAttr(maxLengthEl, "val");
                    if (v != null && Number.isFinite(v) && v > 0 && v <= 10000) {
                        result.maxLength = v;
                    }
                }
                return result;
            }
            const checkBox = globalXmlParser.element(ffData, "checkBox");
            if (checkBox) {
                result.formFieldType = 'checkbox';
                const checkedEl = globalXmlParser.element(checkBox, "checked");
                if (checkedEl) {
                    const v = globalXmlParser.attr(checkedEl, "val");
                    result.checked = v == null ? true : (v === "1" || v === "true");
                }
                else {
                    const defaultEl = globalXmlParser.element(checkBox, "default");
                    if (defaultEl) {
                        const v = globalXmlParser.attr(defaultEl, "val");
                        result.checked = v === "1" || v === "true";
                    }
                    else {
                        result.checked = false;
                    }
                }
                return result;
            }
            const ddList = globalXmlParser.element(ffData, "ddList");
            if (ddList) {
                result.formFieldType = 'dropdown';
                const items = [];
                for (const entry of globalXmlParser.elements(ddList, "listEntry")) {
                    const v = globalXmlParser.attr(entry, "val");
                    if (v != null)
                        items.push(v);
                }
                result.ddItems = items;
                const defaultEl = globalXmlParser.element(ddList, "default");
                if (defaultEl) {
                    const idx = globalXmlParser.intAttr(defaultEl, "val");
                    if (idx != null && idx >= 0 && idx < items.length) {
                        result.ddDefault = idx;
                    }
                }
                return result;
            }
            return result;
        }
        parseSdtControl(sdtPr) {
            for (const el of globalXmlParser.elements(sdtPr)) {
                switch (el.localName) {
                    case "checkbox": {
                        const checkedEl = globalXmlParser.element(el, "checked");
                        const checkedStateEl = globalXmlParser.element(el, "checkedState");
                        const uncheckedStateEl = globalXmlParser.element(el, "uncheckedState");
                        const checkedRaw = checkedEl ? globalXmlParser.attr(checkedEl, "val") : null;
                        const checked = checkedRaw === "1" || checkedRaw === "true";
                        const checkedChar = checkedStateEl ? globalXmlParser.hexAttr(checkedStateEl, "val") : undefined;
                        const uncheckedChar = uncheckedStateEl ? globalXmlParser.hexAttr(uncheckedStateEl, "val") : undefined;
                        const result = { type: "checkbox", checked };
                        if (checkedChar != null)
                            result.checkedChar = checkedChar;
                        if (uncheckedChar != null)
                            result.uncheckedChar = uncheckedChar;
                        return result;
                    }
                    case "dropDownList":
                    case "comboBox": {
                        const items = [];
                        for (const li of globalXmlParser.elements(el, "listItem")) {
                            const displayText = globalXmlParser.attr(li, "displayText");
                            const value = globalXmlParser.attr(li, "value");
                            items.push({
                                displayText: displayText ?? value ?? "",
                                value: value ?? displayText ?? ""
                            });
                        }
                        return { type: "dropdown", items };
                    }
                    case "date":
                    case "sdtDate": {
                        const formatEl = globalXmlParser.element(el, "dateFormat");
                        const fullDateEl = globalXmlParser.element(el, "fullDate");
                        const format = formatEl ? globalXmlParser.attr(formatEl, "val") : null;
                        const fullDateAttr = globalXmlParser.attr(el, "fullDate");
                        const fullDate = fullDateEl ? globalXmlParser.attr(fullDateEl, "val") : fullDateAttr;
                        return {
                            type: "date",
                            format: format ?? undefined,
                            fullDate: fullDate ?? undefined
                        };
                    }
                    case "picture":
                        return { type: "picture" };
                    case "docPartList":
                    case "docPartObj":
                        return { type: "gallery" };
                }
            }
            return null;
        }
        parseInserted(node, parentParser) {
            return {
                type: DomType.Inserted,
                revision: parseRevisionAttrs(node),
                children: parentParser(node)?.children ?? []
            };
        }
        parseDeleted(node, parentParser) {
            return {
                type: DomType.Deleted,
                revision: parseRevisionAttrs(node),
                children: parentParser(node)?.children ?? []
            };
        }
        parseMoveFrom(node, parentParser) {
            return {
                type: DomType.MoveFrom,
                revision: parseRevisionAttrs(node),
                children: parentParser(node)?.children ?? []
            };
        }
        parseMoveTo(node, parentParser) {
            return {
                type: DomType.MoveTo,
                revision: parseRevisionAttrs(node),
                children: parentParser(node)?.children ?? []
            };
        }
        parseAltChunk(node) {
            return { type: DomType.AltChunk, children: [], id: globalXmlParser.attr(node, "id") };
        }
        parseParagraph(node) {
            var result = { type: DomType.Paragraph, children: [] };
            const paraId = node.getAttributeNS("http://schemas.microsoft.com/office/word/2010/wordml", "paraId")
                ?? node.getAttribute("w14:paraId")
                ?? globalXmlParser.attr(node, "paraId");
            if (paraId) {
                result.paraId = paraId;
            }
            for (let el of globalXmlParser.elements(node)) {
                switch (el.localName) {
                    case "pPr":
                        this.parseParagraphProperties(el, result);
                        break;
                    case "r":
                        result.children.push(this.parseRun(el, result));
                        break;
                    case "hyperlink":
                        result.children.push(this.parseHyperlink(el, result));
                        break;
                    case "smartTag":
                        result.children.push(this.parseSmartTag(el, result));
                        break;
                    case "bookmarkStart":
                        result.children.push(parseBookmarkStart(el, globalXmlParser));
                        break;
                    case "bookmarkEnd":
                        result.children.push(parseBookmarkEnd(el, globalXmlParser));
                        break;
                    case "commentRangeStart":
                        result.children.push(new WmlCommentRangeStart(globalXmlParser.attr(el, "id")));
                        break;
                    case "commentRangeEnd":
                        result.children.push(new WmlCommentRangeEnd(globalXmlParser.attr(el, "id")));
                        break;
                    case "oMath":
                    case "oMathPara":
                        result.children.push(this.parseMathElement(el));
                        break;
                    case "sdt":
                        result.children.push(...this.parseSdt(el, e => this.parseParagraph(e).children));
                        break;
                    case "ins":
                        result.children.push(this.parseInserted(el, e => this.parseParagraph(e)));
                        break;
                    case "del":
                        result.children.push(this.parseDeleted(el, e => this.parseParagraph(e)));
                        break;
                    case "moveFrom":
                        result.children.push(this.parseMoveFrom(el, e => this.parseParagraph(e)));
                        break;
                    case "moveTo":
                        result.children.push(this.parseMoveTo(el, e => this.parseParagraph(e)));
                        break;
                    case "fldSimple":
                        result.children.push(this.parseFieldSimple(el, result));
                        break;
                    case "ruby":
                        result.children.push(this.parseRuby(el, result));
                        break;
                }
            }
            return result;
        }
        parseFieldSimple(node, parent) {
            const result = {
                type: DomType.SimpleField,
                instruction: globalXmlParser.attr(node, "instr"),
                lock: globalXmlParser.boolAttr(node, "lock", false),
                dirty: globalXmlParser.boolAttr(node, "dirty", false),
                parent,
                children: [],
            };
            for (const c of globalXmlParser.elements(node)) {
                switch (c.localName) {
                    case "r":
                        result.children.push(this.parseRun(c, result));
                        break;
                    case "hyperlink":
                        result.children.push(this.parseHyperlink(c, result));
                        break;
                    case "fldSimple":
                        result.children.push(this.parseFieldSimple(c, result));
                        break;
                }
            }
            return result;
        }
        parseParagraphProperties(elem, paragraph) {
            this.parseDefaultProperties(elem, paragraph.cssStyle = {}, null, c => {
                if (parseParagraphProperty(c, paragraph, globalXmlParser))
                    return true;
                switch (c.localName) {
                    case "pStyle":
                        paragraph.styleName = globalXmlParser.attr(c, "val");
                        break;
                    case "cnfStyle":
                        paragraph.className = values.classNameOfCnfStyle(c);
                        break;
                    case "framePr":
                        this.parseFrame(c, paragraph);
                        break;
                    case "rPr":
                        for (const rPrChild of globalXmlParser.elements(c)) {
                            if (rPrChild.localName === "ins") {
                                paragraph.paragraphMarkRevisionKind = 'inserted';
                                paragraph.revision = parseRevisionAttrs(rPrChild);
                            }
                            else if (rPrChild.localName === "del") {
                                paragraph.paragraphMarkRevisionKind = 'deleted';
                                paragraph.revision = parseRevisionAttrs(rPrChild);
                            }
                        }
                        break;
                    case "pPrChange":
                        paragraph.formattingRevision = parseFormattingRevision(c);
                        break;
                    default:
                        return false;
                }
                return true;
            });
        }
        parseFrame(node, paragraph) {
            const dropCap = globalXmlParser.attr(node, "dropCap");
            if (dropCap !== "drop" && dropCap !== "margin")
                return;
            const linesRaw = globalXmlParser.intAttr(node, "lines");
            const lines = (Number.isInteger(linesRaw) && linesRaw >= 1 && linesRaw <= 10)
                ? linesRaw
                : 3;
            paragraph.dropCap = dropCap;
            paragraph.dropCapLines = lines;
        }
        parseHyperlink(node, parent) {
            var result = { type: DomType.Hyperlink, parent: parent, children: [] };
            result.anchor = globalXmlParser.attr(node, "anchor");
            result.id = globalXmlParser.attr(node, "id");
            const tooltip = globalXmlParser.attr(node, "tooltip");
            if (tooltip)
                result.tooltip = tooltip;
            const targetFrame = globalXmlParser.attr(node, "tgtFrame");
            if (targetFrame)
                result.targetFrame = targetFrame;
            for (const c of globalXmlParser.elements(node)) {
                switch (c.localName) {
                    case "r":
                        result.children.push(this.parseRun(c, result));
                        break;
                }
            }
            return result;
        }
        parseSmartTag(node, parent) {
            var result = { type: DomType.SmartTag, parent, children: [] };
            var uri = globalXmlParser.attr(node, "uri");
            var element = globalXmlParser.attr(node, "element");
            if (uri)
                result.uri = uri;
            if (element)
                result.element = element;
            for (const c of globalXmlParser.elements(node)) {
                switch (c.localName) {
                    case "r":
                        result.children.push(this.parseRun(c, result));
                        break;
                    case "smartTag":
                        result.children.push(this.parseSmartTag(c, result));
                        break;
                }
            }
            return result;
        }
        parseRun(node, parent) {
            var result = { type: DomType.Run, parent: parent, children: [] };
            for (let c of globalXmlParser.elements(node)) {
                c = this.checkAlternateContent(c);
                switch (c.localName) {
                    case "t":
                        result.children.push({
                            type: DomType.Text,
                            text: c.textContent
                        });
                        break;
                    case "delText":
                        result.children.push({
                            type: DomType.DeletedText,
                            text: c.textContent
                        });
                        break;
                    case "commentReference":
                        result.children.push(new WmlCommentReference(globalXmlParser.attr(c, "id")));
                        break;
                    case "fldSimple":
                        result.children.push(this.parseFieldSimple(c, result));
                        break;
                    case "instrText":
                        result.fieldRun = true;
                        result.children.push({
                            type: DomType.Instruction,
                            text: c.textContent
                        });
                        break;
                    case "fldChar":
                        result.fieldRun = true;
                        const charType = globalXmlParser.attr(c, "fldCharType");
                        const fldChar = {
                            type: DomType.ComplexField,
                            charType,
                            lock: globalXmlParser.boolAttr(c, "lock", false),
                            dirty: globalXmlParser.boolAttr(c, "dirty", false)
                        };
                        if (charType === "begin") {
                            const ffData = globalXmlParser.element(c, "ffData");
                            if (ffData) {
                                fldChar.ffData = this.parseFfData(ffData);
                            }
                        }
                        result.children.push(fldChar);
                        break;
                    case "noBreakHyphen":
                        result.children.push({ type: DomType.NoBreakHyphen });
                        break;
                    case "softHyphen":
                        result.children.push({
                            type: DomType.Text,
                            text: "­"
                        });
                        break;
                    case "br":
                    case "cr":
                        result.children.push({
                            type: DomType.Break,
                            break: c.localName === "cr" ? "textWrapping" : (globalXmlParser.attr(c, "type") || "textWrapping")
                        });
                        break;
                    case "lastRenderedPageBreak":
                        result.children.push({
                            type: DomType.Break,
                            break: "lastRenderedPageBreak"
                        });
                        break;
                    case "sym":
                        result.children.push({
                            type: DomType.Symbol,
                            font: encloseFontFamily(globalXmlParser.attr(c, "font")),
                            char: globalXmlParser.hexAttr(c, "char")
                        });
                        break;
                    case "tab":
                        result.children.push({ type: DomType.Tab });
                        break;
                    case "footnoteReference":
                        result.children.push({
                            type: DomType.FootnoteReference,
                            id: globalXmlParser.attr(c, "id")
                        });
                        break;
                    case "endnoteReference":
                        result.children.push({
                            type: DomType.EndnoteReference,
                            id: globalXmlParser.attr(c, "id")
                        });
                        break;
                    case "drawing":
                        let d = this.parseDrawing(c);
                        if (d)
                            result.children.push(d);
                        break;
                    case "pict":
                        result.children.push(this.parseVmlPicture(c));
                        break;
                    case "object": {
                        const ole = this.parseOleObject(c);
                        if (ole)
                            result.children.push(ole);
                        break;
                    }
                    case "ruby":
                        result.children.push(this.parseRuby(c, result));
                        break;
                    case "rPr":
                        this.parseRunProperties(c, result);
                        break;
                }
            }
            let wrapped = result;
            if (result.bidiOverride) {
                const bidi = {
                    type: DomType.BidiOverride,
                    dir: result.bidiOverride,
                    parent,
                    children: [wrapped]
                };
                wrapped.parent = bidi;
                delete result.bidiOverride;
                wrapped = bidi;
            }
            if (result.fitText) {
                const fit = {
                    type: DomType.FitText,
                    width: result.fitText.width,
                    id: result.fitText.id,
                    parent,
                    children: [wrapped]
                };
                wrapped.parent = fit;
                delete result.fitText;
                wrapped = fit;
            }
            return wrapped;
        }
        parseRuby(node, parent) {
            const result = { type: DomType.Ruby, parent, children: [] };
            for (const c of globalXmlParser.elements(node)) {
                switch (c.localName) {
                    case "rubyPr": {
                        const rubyPr = {};
                        for (const p of globalXmlParser.elements(c)) {
                            const v = globalXmlParser.attr(p, "val");
                            switch (p.localName) {
                                case "rubyAlign":
                                    if (/^(center|distributeLetter|distributeSpace|left|right|rightVertical|start|end)$/.test(v))
                                        rubyPr.rubyAlign = v;
                                    break;
                                case "hps": {
                                    const n = globalXmlParser.intAttr(p, "val");
                                    if (Number.isFinite(n) && n > 0 && n < 1000)
                                        rubyPr.hps = n;
                                    break;
                                }
                                case "hpsBaseText": {
                                    const n = globalXmlParser.intAttr(p, "val");
                                    if (Number.isFinite(n) && n > 0 && n < 1000)
                                        rubyPr.hpsBaseText = n;
                                    break;
                                }
                                case "hpsRaise": {
                                    const n = globalXmlParser.intAttr(p, "val");
                                    if (Number.isFinite(n) && n >= 0 && n < 1000)
                                        rubyPr.hpsRaise = n;
                                    break;
                                }
                                case "lid":
                                    if (v)
                                        rubyPr.lid = v;
                                    break;
                            }
                        }
                        result.rubyPr = rubyPr;
                        break;
                    }
                    case "rubyBase": {
                        const base = { type: DomType.RubyBase, parent: result, children: [] };
                        for (const r of globalXmlParser.elements(c)) {
                            if (r.localName === "r")
                                base.children.push(this.parseRun(r, base));
                        }
                        result.children.push(base);
                        break;
                    }
                    case "rt": {
                        const rt = { type: DomType.RubyText, parent: result, children: [] };
                        for (const r of globalXmlParser.elements(c)) {
                            if (r.localName === "r")
                                rt.children.push(this.parseRun(r, rt));
                        }
                        result.children.push(rt);
                        break;
                    }
                }
            }
            return result;
        }
        parseMathElement(elem) {
            const propsTag = `${elem.localName}Pr`;
            const result = { type: mmlTagMap[elem.localName], children: [] };
            for (const el of globalXmlParser.elements(elem)) {
                const childType = mmlTagMap[el.localName];
                if (childType) {
                    result.children.push(this.parseMathElement(el));
                }
                else if (el.localName == "r") {
                    var run = this.parseRun(el);
                    run.type = DomType.MmlRun;
                    this.stripMathAlignmentMarkers(run);
                    result.children.push(run);
                }
                else if (el.localName == propsTag) {
                    result.props = this.parseMathProperies(el);
                }
            }
            return result;
        }
        stripMathAlignmentMarkers(run) {
            if (!run.children)
                return;
            for (const child of run.children) {
                if (child.type === DomType.Text || child.type === DomType.DeletedText) {
                    const t = child.text;
                    if (t && t.indexOf('&') >= 0) {
                        child.text = t.replace(/&/g, '');
                    }
                }
            }
        }
        parseMathProperies(elem) {
            const result = {};
            for (const el of globalXmlParser.elements(elem)) {
                switch (el.localName) {
                    case "chr":
                        result.char = globalXmlParser.attr(el, "val");
                        break;
                    case "vertJc":
                        result.verticalJustification = globalXmlParser.attr(el, "val");
                        break;
                    case "pos":
                        result.position = globalXmlParser.attr(el, "val");
                        break;
                    case "degHide":
                        result.hideDegree = globalXmlParser.boolAttr(el, "val");
                        break;
                    case "begChr":
                        result.beginChar = globalXmlParser.attr(el, "val");
                        break;
                    case "endChr":
                        result.endChar = globalXmlParser.attr(el, "val");
                        break;
                    case "limLoc":
                        result.limLoc = globalXmlParser.attr(el, "val");
                        break;
                }
            }
            return result;
        }
        parseRunProperties(elem, run) {
            this.parseDefaultProperties(elem, run.cssStyle = {}, null, c => {
                switch (c.localName) {
                    case "rStyle":
                        run.styleName = globalXmlParser.attr(c, "val");
                        break;
                    case "vertAlign":
                        run.verticalAlign = values.valueOfVertAlign(c, true);
                        break;
                    case "rPrChange":
                        run.formattingRevision = parseFormattingRevision(c);
                        break;
                    case "fitText": {
                        const w = globalXmlParser.floatAttr(c, "val");
                        if (Number.isFinite(w) && w > 0) {
                            run.fitText = { width: w, id: globalXmlParser.attr(c, "id") || undefined };
                        }
                        break;
                    }
                    case "bdo": {
                        const raw = globalXmlParser.attr(c, "val");
                        if (raw === "ltr" || raw === "rtl") {
                            run.bidiOverride = raw;
                        }
                        break;
                    }
                    default:
                        return false;
                }
                return true;
            });
        }
        parseVmlPicture(elem) {
            const result = { type: DomType.VmlPicture, children: [] };
            for (const el of globalXmlParser.elements(elem)) {
                const child = parseVmlElement(el, this);
                child && result.children.push(child);
            }
            return result;
        }
        parseOleObject(elem) {
            for (const c of globalXmlParser.elements(elem)) {
                if (c.localName !== "OLEObject")
                    continue;
                return {
                    type: DomType.OleObject,
                    progId: globalXmlParser.attr(c, "ProgID"),
                    shapeId: globalXmlParser.attr(c, "ShapeID"),
                    objectType: globalXmlParser.attr(c, "Type"),
                };
            }
            return { type: DomType.OleObject };
        }
        checkAlternateContent(elem) {
            if (elem.localName != 'AlternateContent')
                return elem;
            var choice = globalXmlParser.element(elem, "Choice");
            if (choice) {
                var requires = globalXmlParser.attr(choice, "Requires");
                var namespaceURI = elem.lookupNamespaceURI(requires);
                if (supportedNamespaceURIs.includes(namespaceURI))
                    return choice.firstElementChild;
            }
            return globalXmlParser.element(elem, "Fallback")?.firstElementChild;
        }
        parseDrawing(node) {
            for (var n of globalXmlParser.elements(node)) {
                switch (n.localName) {
                    case "inline":
                    case "anchor":
                        return this.parseDrawingWrapper(n);
                }
            }
        }
        parseDrawingWrapper(node) {
            var result = { type: DomType.Drawing, children: [], cssStyle: {}, props: {} };
            var isAnchor = node.localName == "anchor";
            (result.props ?? (result.props = {})).isAnchor = isAnchor;
            const EMU_PER_PX = 9525;
            const WRAP_TEXT_ALLOWED = new Set(["bothSides", "left", "right", "largest"]);
            const RELATIVE_FROM_ALLOWED = new Set([
                "margin", "page", "column", "character", "leftMargin", "rightMargin",
                "insideMargin", "outsideMargin", "paragraph", "line", "topMargin",
                "bottomMargin"
            ]);
            const distT = globalXmlParser.floatAttr(node, "distT", 0);
            const distB = globalXmlParser.floatAttr(node, "distB", 0);
            const distL = globalXmlParser.floatAttr(node, "distL", 0);
            const distR = globalXmlParser.floatAttr(node, "distR", 0);
            const marginTopPx = (distT || 0) / EMU_PER_PX;
            const marginBottomPx = (distB || 0) / EMU_PER_PX;
            const marginLeftPx = (distL || 0) / EMU_PER_PX;
            const marginRightPx = (distR || 0) / EMU_PER_PX;
            let wrapType = null;
            let wrapText = null;
            let wrapPolygonPoints = null;
            let simplePos = globalXmlParser.boolAttr(node, "simplePos");
            globalXmlParser.boolAttr(node, "behindDoc");
            let extentCx = 0;
            let extentCy = 0;
            let posX = { relative: "page", align: "left", offset: "0", offsetEmu: 0 };
            let posY = { relative: "page", align: "top", offset: "0", offsetEmu: 0 };
            let docPrDescr = null;
            for (var n of globalXmlParser.elements(node)) {
                switch (n.localName) {
                    case "docPr":
                        docPrDescr = globalXmlParser.attr(n, "descr");
                        break;
                    case "simplePos":
                        if (simplePos) {
                            posX.offsetEmu = globalXmlParser.floatAttr(n, "x", 0) || 0;
                            posY.offsetEmu = globalXmlParser.floatAttr(n, "y", 0) || 0;
                            posX.offset = globalXmlParser.lengthAttr(n, "x", LengthUsage.Emu);
                            posY.offset = globalXmlParser.lengthAttr(n, "y", LengthUsage.Emu);
                        }
                        break;
                    case "extent":
                        extentCx = globalXmlParser.floatAttr(n, "cx", 0) || 0;
                        extentCy = globalXmlParser.floatAttr(n, "cy", 0) || 0;
                        result.cssStyle["width"] = globalXmlParser.lengthAttr(n, "cx", LengthUsage.Emu);
                        result.cssStyle["height"] = globalXmlParser.lengthAttr(n, "cy", LengthUsage.Emu);
                        break;
                    case "positionH":
                    case "positionV":
                        if (!simplePos) {
                            let pos = n.localName == "positionH" ? posX : posY;
                            var alignNode = globalXmlParser.element(n, "align");
                            var offsetNode = globalXmlParser.element(n, "posOffset");
                            pos.relative = globalXmlParser.attr(n, "relativeFrom") ?? pos.relative;
                            if (alignNode)
                                pos.align = alignNode.textContent;
                            if (offsetNode) {
                                pos.offset = convertLength(offsetNode.textContent, LengthUsage.Emu);
                                const parsed = parseFloat(offsetNode.textContent);
                                pos.offsetEmu = Number.isFinite(parsed) ? parsed : 0;
                            }
                        }
                        break;
                    case "wrapTopAndBottom":
                        wrapType = "wrapTopAndBottom";
                        break;
                    case "wrapNone":
                        wrapType = "wrapNone";
                        break;
                    case "wrapSquare":
                        wrapType = "wrapSquare";
                        wrapText = globalXmlParser.attr(n, "wrapText");
                        break;
                    case "wrapTight":
                    case "wrapThrough":
                        wrapType = n.localName;
                        wrapText = globalXmlParser.attr(n, "wrapText");
                        {
                            const polyNode = globalXmlParser.element(n, "wrapPolygon");
                            if (polyNode) {
                                const pts = [];
                                for (const child of globalXmlParser.elements(polyNode)) {
                                    if (child.localName !== "start" && child.localName !== "lineTo")
                                        continue;
                                    const x = globalXmlParser.floatAttr(child, "x", NaN);
                                    const y = globalXmlParser.floatAttr(child, "y", NaN);
                                    if (Number.isFinite(x) && Number.isFinite(y))
                                        pts.push([x, y]);
                                }
                                if (pts.length >= 3)
                                    wrapPolygonPoints = pts;
                            }
                        }
                        break;
                    case "graphic":
                        var g = this.parseGraphic(n);
                        if (g)
                            result.children.push(g);
                        break;
                }
            }
            const safeWrapText = wrapText && WRAP_TEXT_ALLOWED.has(wrapText) ? wrapText : null;
            const safeRelativeH = RELATIVE_FROM_ALLOWED.has(posX.relative) ? posX.relative : "page";
            const floatFromWrapText = (wt, align) => {
                if (wt === "left")
                    return "right";
                if (wt === "right")
                    return "left";
                if (align === "right")
                    return "right";
                return "left";
            };
            const applyMarginsToStyle = () => {
                if (distT)
                    result.cssStyle["margin-top"] = `${marginTopPx.toFixed(2)}px`;
                if (distB)
                    result.cssStyle["margin-bottom"] = `${marginBottomPx.toFixed(2)}px`;
                if (distL)
                    result.cssStyle["margin-left"] = `${marginLeftPx.toFixed(2)}px`;
                if (distR)
                    result.cssStyle["margin-right"] = `${marginRightPx.toFixed(2)}px`;
            };
            const applyShapeMargin = () => {
                const maxDist = Math.max(marginTopPx, marginBottomPx, marginLeftPx, marginRightPx);
                if (maxDist > 0)
                    result.cssStyle["shape-margin"] = `${maxDist.toFixed(2)}px`;
            };
            const buildPolygonCss = () => {
                if (!wrapPolygonPoints || wrapPolygonPoints.length < 3)
                    return null;
                if (extentCx <= 0 || extentCy <= 0)
                    return null;
                const segs = wrapPolygonPoints.map(([x, y]) => {
                    const px = (x / extentCx) * 100;
                    const py = (y / extentCy) * 100;
                    return `${px.toFixed(2)}% ${py.toFixed(2)}%`;
                });
                return `polygon(${segs.join(", ")})`;
            };
            if (wrapType == "wrapTopAndBottom") {
                result.cssStyle['display'] = 'block';
                if (posX.align) {
                    result.cssStyle['text-align'] = posX.align;
                    result.cssStyle['width'] = "100%";
                }
            }
            else if (wrapType == "wrapNone") {
                result.cssStyle['display'] = 'block';
                result.cssStyle['position'] = 'relative';
                result.cssStyle["width"] = "0px";
                result.cssStyle["height"] = "0px";
                if (posX.offset)
                    result.cssStyle["left"] = posX.offset;
                if (posY.offset)
                    result.cssStyle["top"] = posY.offset;
            }
            else if (wrapType == "wrapSquare" || wrapType == "wrapTight" || wrapType == "wrapThrough") {
                const floatSide = floatFromWrapText(safeWrapText, posX.align);
                if (safeRelativeH === "paragraph") {
                    result.cssStyle["position"] = "absolute";
                    const leftPx = (posX.offsetEmu || 0) / EMU_PER_PX;
                    result.cssStyle["left"] = `${leftPx.toFixed(2)}px`;
                    if (posY.offsetEmu) {
                        const topPx = (posY.offsetEmu || 0) / EMU_PER_PX;
                        result.cssStyle["top"] = `${topPx.toFixed(2)}px`;
                    }
                    applyMarginsToStyle();
                }
                else if (safeRelativeH === "column") {
                    result.cssStyle["display"] = "inline-block";
                    result.cssStyle["float"] = floatSide;
                    applyMarginsToStyle();
                    if (floatSide === "left" && distL)
                        result.cssStyle["margin-left"] = `${(-marginLeftPx).toFixed(2)}px`;
                    else if (floatSide === "right" && distR)
                        result.cssStyle["margin-right"] = `${(-marginRightPx).toFixed(2)}px`;
                }
                else {
                    result.cssStyle["float"] = floatSide;
                    applyMarginsToStyle();
                }
                applyShapeMargin();
                if (wrapType === "wrapTight" || wrapType === "wrapThrough") {
                    const polyCss = buildPolygonCss();
                    if (polyCss)
                        result.cssStyle["shape-outside"] = polyCss;
                }
            }
            else if (isAnchor && (posX.align == 'left' || posX.align == 'right')) {
                result.cssStyle["float"] = posX.align;
            }
            if (docPrDescr != null) {
                this.setImageAltText(result, docPrDescr);
            }
            return result;
        }
        setImageAltText(elem, descr) {
            if (elem.type === DomType.Image) {
                elem.altText = descr;
                return;
            }
            if (elem.children) {
                for (const c of elem.children)
                    this.setImageAltText(c, descr);
            }
        }
        parseGraphic(elem) {
            var graphicData = globalXmlParser.element(elem, "graphicData");
            const uri = graphicData ? globalXmlParser.attr(graphicData, "uri") : null;
            const CHARTEX_URI = "http://schemas.microsoft.com/office/drawingml/2014/chartex";
            const SMARTART_URI = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
            if (uri === SMARTART_URI) {
                return this.parseSmartArtReference(elem, graphicData);
            }
            for (let n of globalXmlParser.elements(graphicData)) {
                if (uri === CHARTEX_URI && n.localName === "chart") {
                    return this.parseChartExReference(n);
                }
                switch (n.localName) {
                    case "pic":
                        return this.parsePicture(n);
                    case "wsp":
                        return this.parseDrawingShape(n);
                    case "wgp":
                        return this.parseDrawingShapeGroup(n);
                    case "chart":
                        return this.parseChartReference(n);
                }
            }
            return null;
        }
        parseChartReference(elem) {
            const relId = globalXmlParser.attr(elem, "id");
            return {
                type: DomType.Chart,
                relId: relId ?? "",
            };
        }
        parseChartExReference(elem) {
            const relId = globalXmlParser.attr(elem, "id");
            return {
                type: DomType.ChartEx,
                relId: relId ?? "",
            };
        }
        parseSmartArtReference(graphic, graphicData) {
            const alt = findAncestorByLocalName(graphic, "AlternateContent");
            if (alt) {
                const fallback = globalXmlParser.element(alt, "Fallback");
                if (fallback) {
                    const drawing = globalXmlParser.element(fallback, "drawing");
                    if (drawing) {
                        const r = this.parseDrawing(drawing);
                        if (r)
                            return r;
                    }
                    for (const c of globalXmlParser.elements(fallback)) {
                        if (c.localName === "inline" || c.localName === "anchor") {
                            const r = this.parseDrawingWrapper(c);
                            if (r)
                                return r;
                        }
                    }
                    for (const c of globalXmlParser.elements(fallback)) {
                        if (c.localName === "pic") {
                            return this.parsePicture(c);
                        }
                    }
                }
            }
            const relIds = {};
            if (graphicData) {
                const rel = globalXmlParser.element(graphicData, "relIds");
                if (rel) {
                    const dm = globalXmlParser.attr(rel, "dm");
                    const lo = globalXmlParser.attr(rel, "lo");
                    const qs = globalXmlParser.attr(rel, "qs");
                    const cs = globalXmlParser.attr(rel, "cs");
                    if (dm)
                        relIds.dm = dm;
                    if (lo)
                        relIds.lo = lo;
                    if (qs)
                        relIds.qs = qs;
                    if (cs)
                        relIds.cs = cs;
                }
            }
            return {
                type: DomType.SmartArt,
                children: [],
                cssStyle: {},
                relIds,
            };
        }
        parseXfrm(xfrm) {
            if (!xfrm)
                return undefined;
            const result = { x: 0, y: 0, cx: 0, cy: 0 };
            const rotRaw = globalXmlParser.intAttr(xfrm, "rot", 0);
            if (rotRaw)
                result.rot = rotRaw / 60000;
            for (const n of globalXmlParser.elements(xfrm)) {
                switch (n.localName) {
                    case "off":
                        result.x = globalXmlParser.floatAttr(n, "x", 0) || 0;
                        result.y = globalXmlParser.floatAttr(n, "y", 0) || 0;
                        break;
                    case "ext":
                        result.cx = globalXmlParser.floatAttr(n, "cx", 0) || 0;
                        result.cy = globalXmlParser.floatAttr(n, "cy", 0) || 0;
                        break;
                }
            }
            return result;
        }
        parseShapeFill(spPr) {
            if (!spPr)
                return undefined;
            for (const n of globalXmlParser.elements(spPr)) {
                switch (n.localName) {
                    case "noFill":
                        return { type: "none" };
                    case "solidFill": {
                        const ref = this.parseColor(n);
                        if (ref?.hex)
                            return { type: "solid", color: ref.hex };
                        if (ref?.scheme) {
                            return {
                                type: "gradient",
                                gradient: {
                                    kind: "linear",
                                    stops: [
                                        { pos: 0, colour: ref },
                                        { pos: 1, colour: ref },
                                    ],
                                    angle: 0,
                                },
                            };
                        }
                        return undefined;
                    }
                    case "gradFill": {
                        const grad = this.parseGradientFill(n);
                        if (grad)
                            return { type: "gradient", gradient: grad };
                        return undefined;
                    }
                    case "pattFill": {
                        const patt = this.parsePatternFill(n);
                        if (patt)
                            return { type: "pattern", pattern: patt };
                        return undefined;
                    }
                    case "blipFill":
                        return undefined;
                }
            }
            return undefined;
        }
        parseShapeStroke(spPr) {
            if (!spPr)
                return undefined;
            const ln = globalXmlParser.element(spPr, "ln");
            if (!ln)
                return undefined;
            const result = {};
            const w = globalXmlParser.intAttr(ln, "w", 0);
            if (w > 0)
                result.width = w;
            const solid = globalXmlParser.element(ln, "solidFill");
            if (solid) {
                const ref = this.parseColor(solid);
                if (ref?.hex)
                    result.color = ref.hex;
                else if (ref?.scheme) {
                    const p = DEFAULT_THEME_PALETTE[ref.scheme];
                    const sanitized = sanitizeCssColor(p);
                    if (sanitized)
                        result.color = sanitized;
                }
            }
            return result;
        }
        parseColor(elem) {
            if (!elem)
                return null;
            const srgb = globalXmlParser.element(elem, "srgbClr") ?? (elem.localName === "srgbClr" ? elem : null);
            const scheme = globalXmlParser.element(elem, "schemeClr") ?? (elem.localName === "schemeClr" ? elem : null);
            const sys = globalXmlParser.element(elem, "sysClr") ?? (elem.localName === "sysClr" ? elem : null);
            const ref = {};
            let modsSource = null;
            if (srgb) {
                const val = globalXmlParser.attr(srgb, "val");
                const sanitized = sanitizeCssColor(val);
                if (sanitized)
                    ref.hex = sanitized;
                modsSource = srgb;
            }
            else if (sys) {
                const val = globalXmlParser.attr(sys, "lastClr") ?? globalXmlParser.attr(sys, "val");
                const sanitized = sanitizeCssColor(val);
                if (sanitized)
                    ref.hex = sanitized;
                modsSource = sys;
            }
            else if (scheme) {
                const val = globalXmlParser.attr(scheme, "val");
                if (isAllowedSchemeSlot(val))
                    ref.scheme = val;
                modsSource = scheme;
            }
            if (modsSource) {
                const mods = {};
                let gotMod = false;
                for (const m of globalXmlParser.elements(modsSource)) {
                    const v = globalXmlParser.intAttr(m, "val");
                    if (v == null || !Number.isFinite(v))
                        continue;
                    switch (m.localName) {
                        case "lumMod":
                            mods.lumMod = v;
                            gotMod = true;
                            break;
                        case "lumOff":
                            mods.lumOff = v;
                            gotMod = true;
                            break;
                        case "tint":
                            mods.tint = v;
                            gotMod = true;
                            break;
                        case "shade":
                            mods.shade = v;
                            gotMod = true;
                            break;
                        case "alpha":
                            mods.alpha = v;
                            gotMod = true;
                            break;
                    }
                }
                if (gotMod)
                    ref.mods = mods;
            }
            if (!ref.hex && !ref.scheme)
                return null;
            return ref;
        }
        parseGradientFill(el) {
            const gsLst = globalXmlParser.element(el, "gsLst");
            if (!gsLst)
                return null;
            const stops = [];
            for (const gs of globalXmlParser.elements(gsLst, "gs")) {
                const posRaw = globalXmlParser.intAttr(gs, "pos");
                if (posRaw == null || !Number.isFinite(posRaw))
                    continue;
                const pos = Math.max(0, Math.min(1, posRaw / 100000));
                const colour = this.parseColor(gs);
                if (!colour)
                    continue;
                stops.push({ pos, colour });
            }
            if (stops.length === 0)
                return null;
            const lin = globalXmlParser.element(el, "lin");
            const pathEl = globalXmlParser.element(el, "path");
            if (pathEl) {
                const path = globalXmlParser.attr(pathEl, "path");
                const radialPath = path === 'rect' ? 'rect' : 'circle';
                return { kind: 'radial', stops, path: radialPath };
            }
            let angle = 0;
            if (lin) {
                const ang = globalXmlParser.intAttr(lin, "ang");
                if (ang != null && Number.isFinite(ang)) {
                    angle = (ang / 60000) % 360;
                }
            }
            return { kind: 'linear', stops, angle };
        }
        parsePatternFill(el) {
            const prst = globalXmlParser.attr(el, "prst") ?? "";
            const preset = this.PATTERN_ALLOWLIST.has(prst) ? prst : "";
            const fgEl = globalXmlParser.element(el, "fgClr");
            const bgEl = globalXmlParser.element(el, "bgClr");
            const fg = fgEl ? this.parseColor(fgEl) : null;
            const bg = bgEl ? this.parseColor(bgEl) : null;
            if (!fg && !bg)
                return null;
            const result = { preset };
            if (fg)
                result.fg = fg;
            if (bg)
                result.bg = bg;
            return result;
        }
        parseAdjustments(avLst) {
            if (!avLst)
                return undefined;
            const out = {};
            let any = false;
            for (const gd of globalXmlParser.elements(avLst, "gd")) {
                const name = globalXmlParser.attr(gd, "name");
                if (!name || !this.ADJUSTMENT_NAME_ALLOWLIST.has(name))
                    continue;
                const fmla = globalXmlParser.attr(gd, "fmla");
                if (!fmla)
                    continue;
                const m = /^val\s+(-?\d+(?:\.\d+)?)$/.exec(fmla.trim());
                if (!m)
                    continue;
                const n = Number(m[1]);
                if (!Number.isFinite(n))
                    continue;
                out[name] = n;
                any = true;
            }
            return any ? out : undefined;
        }
        parseEffectList(el) {
            if (!el)
                return undefined;
            const result = {};
            let any = false;
            for (const n of globalXmlParser.elements(el)) {
                switch (n.localName) {
                    case "outerShdw":
                    case "innerShdw": {
                        const blurRad = globalXmlParser.intAttr(n, "blurRad");
                        const dist = globalXmlParser.intAttr(n, "dist");
                        const dirRaw = globalXmlParser.intAttr(n, "dir");
                        const colour = this.parseColor(n);
                        const entry = {};
                        if (blurRad != null && Number.isFinite(blurRad))
                            entry.blurRad = blurRad;
                        if (dist != null && Number.isFinite(dist))
                            entry.dist = dist;
                        if (dirRaw != null && Number.isFinite(dirRaw))
                            entry.dir = (dirRaw / 60000) % 360;
                        if (colour)
                            entry.colour = colour;
                        if (n.localName === "outerShdw")
                            result.outerShadow = entry;
                        else
                            result.innerShadow = entry;
                        any = true;
                        break;
                    }
                    case "softEdge": {
                        const rad = globalXmlParser.intAttr(n, "rad");
                        if (rad != null && Number.isFinite(rad) && rad > 0) {
                            result.softEdge = { rad };
                            any = true;
                        }
                        break;
                    }
                    case "glow": {
                        const rad = globalXmlParser.intAttr(n, "rad");
                        if (rad != null && Number.isFinite(rad) && rad > 0) {
                            const colour = this.parseColor(n);
                            const entry = { rad };
                            if (colour)
                                entry.colour = colour;
                            result.glow = entry;
                            any = true;
                        }
                        break;
                    }
                    case "reflection": {
                        const stA = globalXmlParser.intAttr(n, "stA");
                        const endA = globalXmlParser.intAttr(n, "endA");
                        const dist = globalXmlParser.intAttr(n, "dist");
                        const dirRaw = globalXmlParser.intAttr(n, "dir");
                        const fadeDirRaw = globalXmlParser.intAttr(n, "fadeDir");
                        const stPos = globalXmlParser.intAttr(n, "stPos");
                        const endPos = globalXmlParser.intAttr(n, "endPos");
                        const rotWithShape = globalXmlParser.boolAttr(n, "rotWithShape");
                        const entry = {};
                        if (stA != null && Number.isFinite(stA))
                            entry.stA = stA;
                        if (endA != null && Number.isFinite(endA))
                            entry.endA = endA;
                        if (dist != null && Number.isFinite(dist))
                            entry.dist = dist;
                        if (dirRaw != null && Number.isFinite(dirRaw))
                            entry.dir = (dirRaw / 60000) % 360;
                        if (fadeDirRaw != null && Number.isFinite(fadeDirRaw))
                            entry.fadeDir = (fadeDirRaw / 60000) % 360;
                        if (stPos != null && Number.isFinite(stPos))
                            entry.stPos = stPos;
                        if (endPos != null && Number.isFinite(endPos))
                            entry.endPos = endPos;
                        if (rotWithShape != null)
                            entry.rotWithShape = rotWithShape;
                        result.reflection = entry;
                        any = true;
                        break;
                    }
                }
            }
            return any ? result : undefined;
        }
        parseBodyPr(elem) {
            if (!elem)
                return undefined;
            const result = {};
            const lIns = globalXmlParser.intAttr(elem, "lIns");
            const tIns = globalXmlParser.intAttr(elem, "tIns");
            const rIns = globalXmlParser.intAttr(elem, "rIns");
            const bIns = globalXmlParser.intAttr(elem, "bIns");
            if (lIns != null)
                result.lIns = lIns;
            if (tIns != null)
                result.tIns = tIns;
            if (rIns != null)
                result.rIns = rIns;
            if (bIns != null)
                result.bIns = bIns;
            return result;
        }
        parseCustGeom(custGeom) {
            if (!custGeom)
                return undefined;
            const paths = [];
            const pathLst = globalXmlParser.element(custGeom, "pathLst");
            if (!pathLst)
                return undefined;
            for (const pathEl of globalXmlParser.elements(pathLst, "path")) {
                const w = this.safeNum(globalXmlParser.intAttr(pathEl, "w", 0));
                const h = this.safeNum(globalXmlParser.intAttr(pathEl, "h", 0));
                if (w <= 0 || h <= 0)
                    continue;
                let d = "";
                let penX = 0;
                let penY = 0;
                for (const cmd of globalXmlParser.elements(pathEl)) {
                    switch (cmd.localName) {
                        case "moveTo": {
                            const pt = globalXmlParser.element(cmd, "pt");
                            if (!pt)
                                break;
                            const x = this.safeNum(globalXmlParser.intAttr(pt, "x", 0));
                            const y = this.safeNum(globalXmlParser.intAttr(pt, "y", 0));
                            d += `M ${x} ${y} `;
                            penX = x;
                            penY = y;
                            break;
                        }
                        case "lnTo": {
                            const pt = globalXmlParser.element(cmd, "pt");
                            if (!pt)
                                break;
                            const x = this.safeNum(globalXmlParser.intAttr(pt, "x", 0));
                            const y = this.safeNum(globalXmlParser.intAttr(pt, "y", 0));
                            d += `L ${x} ${y} `;
                            penX = x;
                            penY = y;
                            break;
                        }
                        case "cubicBezTo": {
                            const pts = globalXmlParser.elements(cmd, "pt");
                            if (pts.length < 3)
                                break;
                            const x1 = this.safeNum(globalXmlParser.intAttr(pts[0], "x", 0));
                            const y1 = this.safeNum(globalXmlParser.intAttr(pts[0], "y", 0));
                            const x2 = this.safeNum(globalXmlParser.intAttr(pts[1], "x", 0));
                            const y2 = this.safeNum(globalXmlParser.intAttr(pts[1], "y", 0));
                            const x = this.safeNum(globalXmlParser.intAttr(pts[2], "x", 0));
                            const y = this.safeNum(globalXmlParser.intAttr(pts[2], "y", 0));
                            d += `C ${x1} ${y1} ${x2} ${y2} ${x} ${y} `;
                            penX = x;
                            penY = y;
                            break;
                        }
                        case "quadBezTo": {
                            const pts = globalXmlParser.elements(cmd, "pt");
                            if (pts.length < 2)
                                break;
                            const x1 = this.safeNum(globalXmlParser.intAttr(pts[0], "x", 0));
                            const y1 = this.safeNum(globalXmlParser.intAttr(pts[0], "y", 0));
                            const x = this.safeNum(globalXmlParser.intAttr(pts[1], "x", 0));
                            const y = this.safeNum(globalXmlParser.intAttr(pts[1], "y", 0));
                            d += `Q ${x1} ${y1} ${x} ${y} `;
                            penX = x;
                            penY = y;
                            break;
                        }
                        case "arcTo": {
                            const wR = this.safeNum(globalXmlParser.intAttr(cmd, "wR", 0));
                            const hR = this.safeNum(globalXmlParser.intAttr(cmd, "hR", 0));
                            const stAng = this.safeNum(globalXmlParser.intAttr(cmd, "stAng", 0));
                            const swAng = this.safeNum(globalXmlParser.intAttr(cmd, "swAng", 0));
                            if (wR === 0 || hR === 0)
                                break;
                            const stRad = (stAng / 60000) * (Math.PI / 180);
                            const swRad = (swAng / 60000) * (Math.PI / 180);
                            const cx = penX - wR * Math.cos(stRad);
                            const cy = penY - hR * Math.sin(stRad);
                            const endAng = stRad + swRad;
                            const endX = cx + wR * Math.cos(endAng);
                            const endY = cy + hR * Math.sin(endAng);
                            const largeArc = Math.abs(swRad) > Math.PI ? 1 : 0;
                            const sweep = swRad >= 0 ? 1 : 0;
                            d += `A ${wR} ${hR} 0 ${largeArc} ${sweep} ${endX} ${endY} `;
                            penX = endX;
                            penY = endY;
                            break;
                        }
                        case "close": {
                            d += "Z ";
                            break;
                        }
                    }
                }
                d = d.trim();
                if (d) {
                    paths.push({ w, h, d });
                }
            }
            return { paths };
        }
        safeNum(n) {
            return typeof n === "number" && Number.isFinite(n) ? n : 0;
        }
        parseDrawingShape(elem) {
            const result = {
                type: DomType.DrawingShape,
                children: [],
            };
            const spPr = globalXmlParser.element(elem, "spPr");
            if (spPr) {
                const xfrm = globalXmlParser.element(spPr, "xfrm");
                result.xfrm = this.parseXfrm(xfrm) ?? { x: 0, y: 0, cx: 0, cy: 0 };
                const prstGeom = globalXmlParser.element(spPr, "prstGeom");
                if (prstGeom) {
                    const prst = globalXmlParser.attr(prstGeom, "prst");
                    result.presetGeometry = this.PRESET_GEOMETRY_ALLOWLIST.has(prst) ? prst : "rect";
                    const avLst = globalXmlParser.element(prstGeom, "avLst");
                    const adjustments = this.parseAdjustments(avLst);
                    if (adjustments)
                        result.presetAdjustments = adjustments;
                }
                else if (globalXmlParser.element(spPr, "custGeom")) {
                    result.presetGeometry = "rect";
                    result.hasCustomGeometry = true;
                    const custGeom = this.parseCustGeom(globalXmlParser.element(spPr, "custGeom"));
                    if (custGeom && custGeom.paths.length > 0) {
                        result.custGeom = custGeom;
                    }
                }
                else {
                    result.presetGeometry = "rect";
                }
                result.fill = this.parseShapeFill(spPr);
                result.stroke = this.parseShapeStroke(spPr);
                const effects = this.parseEffectList(globalXmlParser.element(spPr, "effectLst"));
                if (effects)
                    result.effects = effects;
            }
            else {
                result.presetGeometry = "rect";
                result.xfrm = { x: 0, y: 0, cx: 0, cy: 0 };
            }
            const bodyPr = globalXmlParser.element(elem, "bodyPr");
            if (bodyPr) {
                result.bodyPr = this.parseBodyPr(bodyPr);
            }
            const txbx = globalXmlParser.element(elem, "txbx");
            const txbxContent = txbx ? globalXmlParser.element(txbx, "txbxContent") : null;
            if (txbxContent) {
                result.txbxParagraphs = this.parseBodyElements(txbxContent);
            }
            return result;
        }
        parseDrawingShapeGroup(elem) {
            const result = {
                type: DomType.DrawingGroup,
                children: [],
            };
            const grpSpPr = globalXmlParser.element(elem, "grpSpPr");
            if (grpSpPr) {
                const xfrm = globalXmlParser.element(grpSpPr, "xfrm");
                if (xfrm) {
                    result.xfrm = this.parseXfrm(xfrm) ?? { x: 0, y: 0, cx: 0, cy: 0 };
                    const chOff = globalXmlParser.element(xfrm, "chOff");
                    const chExt = globalXmlParser.element(xfrm, "chExt");
                    result.childOffset = {
                        x: chOff ? (globalXmlParser.floatAttr(chOff, "x", 0) || 0) : 0,
                        y: chOff ? (globalXmlParser.floatAttr(chOff, "y", 0) || 0) : 0,
                        cx: chExt ? (globalXmlParser.floatAttr(chExt, "cx", 0) || 0) : (result.xfrm?.cx ?? 0),
                        cy: chExt ? (globalXmlParser.floatAttr(chExt, "cy", 0) || 0) : (result.xfrm?.cy ?? 0),
                    };
                }
            }
            for (const n of globalXmlParser.elements(elem)) {
                switch (n.localName) {
                    case "wsp":
                        result.children.push(this.parseDrawingShape(n));
                        break;
                    case "wgp":
                        result.children.push(this.parseDrawingShapeGroup(n));
                        break;
                    case "pic":
                        result.children.push(this.parsePicture(n));
                        break;
                }
            }
            return result;
        }
        parsePicture(elem) {
            var result = { type: DomType.Image, src: "", cssStyle: {} };
            var blipFill = globalXmlParser.element(elem, "blipFill");
            var blip = globalXmlParser.element(blipFill, "blip");
            var srcRect = globalXmlParser.element(blipFill, "srcRect");
            result.src = globalXmlParser.attr(blip, "embed");
            const nvPicPr = globalXmlParser.element(elem, "nvPicPr");
            const cNvPr = nvPicPr ? globalXmlParser.element(nvPicPr, "cNvPr") : null;
            const picDescr = cNvPr ? globalXmlParser.attr(cNvPr, "descr") : null;
            const blipDescr = blip ? globalXmlParser.attr(blip, "descr") : null;
            if (picDescr != null)
                result.altText = picDescr;
            else if (blipDescr != null)
                result.altText = blipDescr;
            if (srcRect) {
                result.srcRect = [
                    globalXmlParser.intAttr(srcRect, "l", 0) / 100000,
                    globalXmlParser.intAttr(srcRect, "t", 0) / 100000,
                    globalXmlParser.intAttr(srcRect, "r", 0) / 100000,
                    globalXmlParser.intAttr(srcRect, "b", 0) / 100000,
                ];
            }
            var spPr = globalXmlParser.element(elem, "spPr");
            var xfrm = globalXmlParser.element(spPr, "xfrm");
            result.cssStyle["position"] = "relative";
            if (xfrm) {
                result.rotation = globalXmlParser.intAttr(xfrm, "rot", 0) / 60000;
                for (var n of globalXmlParser.elements(xfrm)) {
                    switch (n.localName) {
                        case "ext":
                            result.cssStyle["width"] = globalXmlParser.lengthAttr(n, "cx", LengthUsage.Emu);
                            result.cssStyle["height"] = globalXmlParser.lengthAttr(n, "cy", LengthUsage.Emu);
                            break;
                        case "off":
                            result.cssStyle["left"] = globalXmlParser.lengthAttr(n, "x", LengthUsage.Emu);
                            result.cssStyle["top"] = globalXmlParser.lengthAttr(n, "y", LengthUsage.Emu);
                            break;
                    }
                }
            }
            return result;
        }
        parseTable(node) {
            var result = { type: DomType.Table, children: [] };
            for (const c of globalXmlParser.elements(node)) {
                switch (c.localName) {
                    case "tr":
                        result.children.push(this.parseTableRow(c));
                        break;
                    case "tblGrid":
                        result.columns = this.parseTableColumns(c);
                        break;
                    case "tblPr":
                        this.parseTableProperties(c, result);
                        break;
                }
            }
            return result;
        }
        parseTableColumns(node) {
            var result = [];
            for (const n of globalXmlParser.elements(node)) {
                switch (n.localName) {
                    case "gridCol":
                        result.push({ width: globalXmlParser.lengthAttr(n, "w") });
                        break;
                }
            }
            return result;
        }
        parseTableProperties(elem, table) {
            table.cssStyle = {};
            table.cellStyle = {};
            this.parseDefaultProperties(elem, table.cssStyle, table.cellStyle, c => {
                switch (c.localName) {
                    case "tblStyle":
                        table.styleName = globalXmlParser.attr(c, "val");
                        break;
                    case "tblLook":
                        table.className = values.classNameOftblLook(c);
                        {
                            const val = globalXmlParser.hexAttr(c, "val", 0);
                            if (globalXmlParser.boolAttr(c, "firstRow") || (val & 0x0020)) {
                                table.firstRowIsHeader = true;
                            }
                        }
                        break;
                    case "tblpPr":
                        this.parseTablePosition(c, table);
                        break;
                    case "tblStyleColBandSize":
                        table.colBandSize = globalXmlParser.intAttr(c, "val");
                        break;
                    case "tblStyleRowBandSize":
                        table.rowBandSize = globalXmlParser.intAttr(c, "val");
                        break;
                    case "hidden":
                        table.cssStyle["display"] = "none";
                        break;
                    default:
                        return false;
                }
                return true;
            });
            switch (table.cssStyle["text-align"]) {
                case "center":
                    delete table.cssStyle["text-align"];
                    table.cssStyle["margin-left"] = "auto";
                    table.cssStyle["margin-right"] = "auto";
                    break;
                case "right":
                    delete table.cssStyle["text-align"];
                    table.cssStyle["margin-left"] = "auto";
                    break;
            }
        }
        parseTablePosition(node, table) {
            var topFromText = globalXmlParser.lengthAttr(node, "topFromText");
            var bottomFromText = globalXmlParser.lengthAttr(node, "bottomFromText");
            var rightFromText = globalXmlParser.lengthAttr(node, "rightFromText");
            var leftFromText = globalXmlParser.lengthAttr(node, "leftFromText");
            const horzAnchor = globalXmlParser.attr(node, "horzAnchor");
            const vertAnchor = globalXmlParser.attr(node, "vertAnchor");
            const tblpXSpec = globalXmlParser.attr(node, "tblpXSpec");
            const tblpYSpec = globalXmlParser.attr(node, "tblpYSpec");
            const tblpX = globalXmlParser.lengthAttr(node, "tblpX");
            const tblpY = globalXmlParser.lengthAttr(node, "tblpY");
            const pageAnchored = horzAnchor === "page" || vertAnchor === "page";
            if (pageAnchored) {
                table.cssStyle["position"] = "absolute";
                if (tblpX)
                    table.cssStyle["left"] = tblpX;
                if (tblpY)
                    table.cssStyle["top"] = tblpY;
            }
            else {
                table.cssStyle["float"] = "left";
            }
            table.cssStyle["margin-bottom"] = values.addSize(table.cssStyle["margin-bottom"], bottomFromText);
            table.cssStyle["margin-left"] = values.addSize(table.cssStyle["margin-left"], leftFromText);
            table.cssStyle["margin-right"] = values.addSize(table.cssStyle["margin-right"], rightFromText);
            table.cssStyle["margin-top"] = values.addSize(table.cssStyle["margin-top"], topFromText);
            if (tblpXSpec === "center") {
                table.cssStyle["margin-left"] = "auto";
                table.cssStyle["margin-right"] = "auto";
            }
            else if (tblpXSpec === "right") {
                table.cssStyle["margin-left"] = "auto";
            }
            if (tblpYSpec) {
                table.cssStyle["$tblp-y-spec"] = tblpYSpec;
            }
        }
        parseTableRow(node) {
            var result = { type: DomType.Row, children: [] };
            for (const c of globalXmlParser.elements(node)) {
                switch (c.localName) {
                    case "tc":
                        result.children.push(this.parseTableCell(c));
                        break;
                    case "bookmarkStart":
                        result.children.push(parseBookmarkStart(c, globalXmlParser));
                        break;
                    case "bookmarkEnd":
                        result.children.push(parseBookmarkEnd(c, globalXmlParser));
                        break;
                    case "trPr":
                    case "tblPrEx":
                        this.parseTableRowProperties(c, result);
                        break;
                }
            }
            return result;
        }
        parseTableRowProperties(elem, row) {
            row.cssStyle = this.parseDefaultProperties(elem, {}, null, c => {
                switch (c.localName) {
                    case "cnfStyle":
                        row.className = values.classNameOfCnfStyle(c);
                        break;
                    case "tblHeader":
                        row.isHeader = globalXmlParser.boolAttr(c, "val", true);
                        break;
                    case "cantSplit":
                        row.cantSplit = globalXmlParser.boolAttr(c, "val", true);
                        break;
                    case "gridBefore":
                        row.gridBefore = globalXmlParser.intAttr(c, "val");
                        break;
                    case "gridAfter":
                        row.gridAfter = globalXmlParser.intAttr(c, "val");
                        break;
                    case "ins":
                        row.revision = parseRevisionAttrs(c);
                        row.rowRevisionKind = "inserted";
                        break;
                    case "del":
                        row.revision = parseRevisionAttrs(c);
                        row.rowRevisionKind = "deleted";
                        break;
                    case "trPrChange":
                        row.formattingRevision = parseFormattingRevision(c);
                        break;
                    default:
                        return false;
                }
                return true;
            });
        }
        parseTableCell(node) {
            var result = { type: DomType.Cell, children: [] };
            for (const c of globalXmlParser.elements(node)) {
                switch (c.localName) {
                    case "tbl":
                        result.children.push(this.parseTable(c));
                        break;
                    case "p":
                        result.children.push(this.parseParagraph(c));
                        break;
                    case "tcPr":
                        this.parseTableCellProperties(c, result);
                        break;
                }
            }
            return result;
        }
        parseTableCellProperties(elem, cell) {
            cell.cssStyle = this.parseDefaultProperties(elem, {}, null, c => {
                switch (c.localName) {
                    case "gridSpan":
                        cell.span = globalXmlParser.intAttr(c, "val", null);
                        break;
                    case "vMerge":
                        cell.verticalMerge = globalXmlParser.attr(c, "val") ?? "continue";
                        break;
                    case "cnfStyle":
                        cell.className = values.classNameOfCnfStyle(c);
                        break;
                    default:
                        return false;
                }
                return true;
            });
            this.parseTableCellVerticalText(elem, cell);
        }
        parseTableCellVerticalText(elem, cell) {
            const directionMap = {
                "btLr": {
                    writingMode: "vertical-rl",
                    transform: "rotate(180deg)"
                },
                "lrTb": {
                    writingMode: "vertical-lr",
                    transform: "none"
                },
                "tbRl": {
                    writingMode: "vertical-rl",
                    transform: "none"
                }
            };
            for (const c of globalXmlParser.elements(elem)) {
                if (c.localName === "textDirection") {
                    const direction = globalXmlParser.attr(c, "val");
                    const style = directionMap[direction] || { writingMode: "horizontal-tb" };
                    cell.cssStyle["writing-mode"] = style.writingMode;
                    cell.cssStyle["transform"] = style.transform;
                }
            }
        }
        parseDefaultProperties(elem, style = null, childStyle = null, handler = null) {
            style = style || {};
            for (const c of globalXmlParser.elements(elem)) {
                if (handler?.(c))
                    continue;
                switch (c.localName) {
                    case "jc":
                        style["text-align"] = values.valueOfJc(c);
                        break;
                    case "textAlignment":
                        style["vertical-align"] = values.valueOfTextAlignment(c);
                        break;
                    case "color": {
                        style["color"] = xmlUtil.colorAttr(c, "val", null, autos.color);
                        const tref = xmlUtil.themeColorReference(c, "themeColor", "val");
                        if (tref)
                            style["$themeColor-color"] = tref;
                        break;
                    }
                    case "sz":
                        style["font-size"] = style["min-height"] = globalXmlParser.lengthAttr(c, "val", LengthUsage.FontSize);
                        break;
                    case "shd":
                        values.applyShd(c, style);
                        break;
                    case "highlight": {
                        style["background-color"] = xmlUtil.colorAttr(c, "val", null, autos.highlight);
                        const tref = xmlUtil.themeColorReference(c, "themeColor", "val");
                        if (tref)
                            style["$themeColor-background-color"] = tref;
                        break;
                    }
                    case "vertAlign":
                        break;
                    case "position":
                        style.verticalAlign = globalXmlParser.lengthAttr(c, "val", LengthUsage.FontSize);
                        break;
                    case "tcW":
                        if (this.options.ignoreWidth)
                            break;
                    case "tblW":
                        style["width"] = values.valueOfSize(c, "w");
                        break;
                    case "trHeight":
                        this.parseTrHeight(c, style);
                        break;
                    case "strike":
                        style["text-decoration"] = globalXmlParser.boolAttr(c, "val", true) ? "line-through" : "none";
                        break;
                    case "dstrike":
                        style["text-decoration"] = globalXmlParser.boolAttr(c, "val", true) ? "line-through double" : "none";
                        break;
                    case "b":
                        style["font-weight"] = globalXmlParser.boolAttr(c, "val", true) ? "bold" : "normal";
                        break;
                    case "i":
                        style["font-style"] = globalXmlParser.boolAttr(c, "val", true) ? "italic" : "normal";
                        break;
                    case "caps":
                        style["text-transform"] = globalXmlParser.boolAttr(c, "val", true) ? "uppercase" : "none";
                        break;
                    case "smallCaps":
                        style["font-variant"] = globalXmlParser.boolAttr(c, "val", true) ? "small-caps" : "none";
                        break;
                    case "u":
                        this.parseUnderline(c, style);
                        break;
                    case "ind":
                    case "tblInd":
                        this.parseIndentation(c, style);
                        break;
                    case "rFonts":
                        this.parseFont(c, style);
                        break;
                    case "tblBorders":
                        this.parseBorderProperties(c, childStyle || style);
                        break;
                    case "tblCellSpacing":
                        style["border-spacing"] = values.valueOfMargin(c);
                        style["border-collapse"] = "separate";
                        break;
                    case "pBdr":
                        this.parseBorderProperties(c, style);
                        break;
                    case "bdr": {
                        style["border"] = values.valueOfBorder(c);
                        const tref = values.themeRefOfBorder(c);
                        if (tref)
                            style["$themeColor-border"] = tref;
                        break;
                    }
                    case "tcBorders":
                        this.parseBorderProperties(c, style);
                        break;
                    case "vanish":
                    case "specVanish":
                        if (globalXmlParser.boolAttr(c, "val", true))
                            style["display"] = "none";
                        break;
                    case "kern":
                        if (globalXmlParser.boolAttr(c, "val", true))
                            style["font-kerning"] = "normal";
                        break;
                    case "w":
                        {
                            const pct = globalXmlParser.intAttr(c, "val");
                            if (pct != null)
                                style["font-stretch"] = `${pct}%`;
                        }
                        break;
                    case "emboss":
                        if (globalXmlParser.boolAttr(c, "val", true))
                            style["text-shadow"] = "1px 1px 1px #fff, -1px -1px 1px #000";
                        break;
                    case "imprint":
                        if (globalXmlParser.boolAttr(c, "val", true))
                            style["text-shadow"] = "1px 1px 1px #000, -1px -1px 1px #fff";
                        break;
                    case "outline":
                        if (globalXmlParser.boolAttr(c, "val", true)) {
                            style["-webkit-text-stroke"] = "1px currentColor";
                            style["color"] = "transparent";
                        }
                        break;
                    case "shadow":
                        if (globalXmlParser.boolAttr(c, "val", true))
                            style["text-shadow"] = "2px 2px 2px rgba(0,0,0,0.5)";
                        break;
                    case "noWrap":
                        if (globalXmlParser.boolAttr(c, "val", true))
                            style["white-space"] = "nowrap";
                        break;
                    case "tblCellMar":
                    case "tcMar":
                        this.parseMarginProperties(c, childStyle || style);
                        break;
                    case "tblLayout":
                        style["table-layout"] = values.valueOfTblLayout(c);
                        break;
                    case "vAlign":
                        style["vertical-align"] = values.valueOfTextAlignment(c);
                        break;
                    case "spacing":
                        if (elem.localName == "pPr") {
                            this.parseSpacing(c, style);
                        }
                        else if (elem.localName == "rPr") {
                            style["letter-spacing"] = globalXmlParser.lengthAttr(c, "val", LengthUsage.Dxa);
                        }
                        break;
                    case "wordWrap":
                        if (globalXmlParser.boolAttr(c, "val"))
                            style["overflow-wrap"] = "break-word";
                        break;
                    case "suppressAutoHyphens":
                        style["hyphens"] = globalXmlParser.boolAttr(c, "val", true) ? "none" : "auto";
                        break;
                    case "lang":
                        {
                            const langVal = globalXmlParser.attr(c, "val")
                                || globalXmlParser.attr(c, "eastAsia")
                                || globalXmlParser.attr(c, "bidi");
                            if (langVal)
                                style["$lang"] = langVal;
                        }
                        break;
                    case "rtl":
                    case "bidi":
                        if (globalXmlParser.boolAttr(c, "val", true))
                            style["direction"] = "rtl";
                        break;
                    case "cs":
                        break;
                    case "bCs":
                        style["font-weight"] = globalXmlParser.boolAttr(c, "val", true) ? "bold" : "normal";
                        break;
                    case "iCs":
                        style["font-style"] = globalXmlParser.boolAttr(c, "val", true) ? "italic" : "normal";
                        break;
                    case "szCs": {
                        const csSize = globalXmlParser.lengthAttr(c, "val", LengthUsage.FontSize);
                        if (csSize) {
                            style["$cs-font-size"] = csSize;
                            if (!style["font-size"])
                                style["font-size"] = csSize;
                        }
                        break;
                    }
                    case "em": {
                        const raw = globalXmlParser.attr(c, "val");
                        switch (raw) {
                            case "dot":
                                style["text-emphasis"] = "filled dot";
                                break;
                            case "comma":
                                style["text-emphasis"] = "filled sesame";
                                break;
                            case "circle":
                                style["text-emphasis"] = "filled circle";
                                break;
                            case "underDot":
                                style["text-emphasis"] = "filled dot";
                                style["text-emphasis-position"] = "under";
                                break;
                        }
                        break;
                    }
                    case "tabs":
                    case "outlineLvl":
                    case "contextualSpacing":
                    case "tblStyleColBandSize":
                    case "tblStyleRowBandSize":
                    case "webHidden":
                    case "pageBreakBefore":
                    case "suppressLineNumbers":
                    case "keepLines":
                    case "keepNext":
                    case "widowControl":
                    case "noProof":
                        break;
                    default:
                        if (this.options.debug)
                            console.warn(`DOCX: Unknown document element: ${elem.localName}.${c.localName}`);
                        break;
                }
            }
            return style;
        }
        parseUnderline(node, style) {
            var val = globalXmlParser.attr(node, "val");
            if (val == null)
                return;
            switch (val) {
                case "dash":
                case "dashDotDotHeavy":
                case "dashDotHeavy":
                case "dashedHeavy":
                case "dashLong":
                case "dashLongHeavy":
                case "dotDash":
                case "dotDotDash":
                    style["text-decoration"] = "underline dashed";
                    break;
                case "dotted":
                case "dottedHeavy":
                    style["text-decoration"] = "underline dotted";
                    break;
                case "double":
                    style["text-decoration"] = "underline double";
                    break;
                case "single":
                case "thick":
                    style["text-decoration"] = "underline";
                    break;
                case "wave":
                case "wavyDouble":
                case "wavyHeavy":
                    style["text-decoration"] = "underline wavy";
                    break;
                case "words":
                    style["text-decoration"] = "underline";
                    break;
                case "none":
                    style["text-decoration"] = "none";
                    break;
            }
            var col = xmlUtil.colorAttr(node, "color");
            if (col)
                style["text-decoration-color"] = col;
        }
        parseFont(node, style) {
            var ascii = globalXmlParser.attr(node, "ascii");
            var asciiTheme = values.themeValue(node, "asciiTheme");
            var eastAsia = globalXmlParser.attr(node, "eastAsia");
            var fonts = [ascii, asciiTheme, eastAsia].filter(x => x).map(x => encloseFontFamily(x));
            if (fonts.length > 0)
                style["font-family"] = [...new Set(fonts)].join(', ');
        }
        parseIndentation(node, style) {
            var firstLine = globalXmlParser.lengthAttr(node, "firstLine");
            var firstLineChars = globalXmlParser.intAttr(node, "firstLineChars", null);
            var hanging = globalXmlParser.lengthAttr(node, "hanging");
            var hangingChars = globalXmlParser.intAttr(node, "hangingChars", null);
            var left = globalXmlParser.lengthAttr(node, "left");
            var leftChars = globalXmlParser.intAttr(node, "leftChars", null);
            var start = globalXmlParser.lengthAttr(node, "start");
            var startChars = globalXmlParser.intAttr(node, "startChars", null);
            var right = globalXmlParser.lengthAttr(node, "right");
            var rightChars = globalXmlParser.intAttr(node, "rightChars", null);
            var end = globalXmlParser.lengthAttr(node, "end");
            var endChars = globalXmlParser.intAttr(node, "endChars", null);
            if (firstLine)
                style["text-indent"] = firstLine;
            if (firstLineChars != null)
                style["text-indent"] = `${firstLineChars / 100}em`;
            if (hanging)
                style["text-indent"] = `-${hanging}`;
            if (hangingChars != null)
                style["text-indent"] = `-${hangingChars / 100}em`;
            if (left || start)
                style["margin-inline-start"] = left || start;
            if (leftChars != null || startChars != null)
                style["margin-inline-start"] = `${(leftChars ?? startChars) / 100}em`;
            if (right || end)
                style["margin-inline-end"] = right || end;
            if (rightChars != null || endChars != null)
                style["margin-inline-end"] = `${(rightChars ?? endChars) / 100}em`;
        }
        parseSpacing(node, style) {
            var before = globalXmlParser.lengthAttr(node, "before");
            var after = globalXmlParser.lengthAttr(node, "after");
            var line = globalXmlParser.intAttr(node, "line", null);
            var lineRule = globalXmlParser.attr(node, "lineRule");
            if (before)
                style["margin-top"] = before;
            if (after)
                style["margin-bottom"] = after;
            if (line !== null) {
                switch (lineRule) {
                    case "auto":
                        style["line-height"] = `${(line / 240).toFixed(2)}`;
                        break;
                    case "atLeast":
                        style["line-height"] = `calc(100% + ${line / 20}pt)`;
                        break;
                    default:
                        style["line-height"] = style["min-height"] = `${line / 20}pt`;
                        break;
                }
            }
        }
        parseMarginProperties(node, output) {
            for (const c of globalXmlParser.elements(node)) {
                switch (c.localName) {
                    case "left":
                        output["padding-left"] = values.valueOfMargin(c);
                        break;
                    case "right":
                        output["padding-right"] = values.valueOfMargin(c);
                        break;
                    case "top":
                        output["padding-top"] = values.valueOfMargin(c);
                        break;
                    case "bottom":
                        output["padding-bottom"] = values.valueOfMargin(c);
                        break;
                }
            }
        }
        parseTrHeight(node, output) {
            switch (globalXmlParser.attr(node, "hRule")) {
                case "exact":
                    output["height"] = globalXmlParser.lengthAttr(node, "val");
                    break;
                case "atLeast":
                default:
                    output["height"] = globalXmlParser.lengthAttr(node, "val");
                    break;
            }
        }
        parseBorderProperties(node, output) {
            const setBorder = (prop, c) => {
                output[prop] = values.valueOfBorder(c);
                const tref = values.themeRefOfBorder(c);
                if (tref)
                    output[`$themeColor-${prop}`] = tref;
            };
            for (const c of globalXmlParser.elements(node)) {
                switch (c.localName) {
                    case "start":
                    case "left":
                        setBorder("border-left", c);
                        break;
                    case "end":
                    case "right":
                        setBorder("border-right", c);
                        break;
                    case "top":
                        setBorder("border-top", c);
                        break;
                    case "bottom":
                        setBorder("border-bottom", c);
                        break;
                    case "tl2br":
                        output["$diag-tlbr"] = values.valueOfBorder(c);
                        break;
                    case "tr2bl":
                        output["$diag-trbl"] = values.valueOfBorder(c);
                        break;
                }
            }
        }
    }
    const knownColors = ['black', 'blue', 'cyan', 'darkBlue', 'darkCyan', 'darkGray', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'green', 'lightGray', 'magenta', 'none', 'red', 'white', 'yellow'];
    class xmlUtil {
        static colorAttr(node, attrName, defValue = null, autoColor = 'black', themeAttrName = "themeColor") {
            var v = globalXmlParser.attr(node, attrName);
            if (v) {
                if (v == "auto") {
                    return autoColor;
                }
                else if (knownColors.includes(v)) {
                    return v;
                }
                return `#${v}`;
            }
            var themeColor = globalXmlParser.attr(node, themeAttrName);
            return themeColor ? `var(--docx-${themeColor}-color)` : defValue;
        }
        static themeColorReference(node, themeAttrName = "themeColor", literalAttrName) {
            if (literalAttrName) {
                const literal = globalXmlParser.attr(node, literalAttrName);
                if (literal && literal !== "auto")
                    return null;
            }
            const slot = globalXmlParser.attr(node, themeAttrName);
            if (!slot)
                return null;
            const tint = globalXmlParser.attr(node, "themeTint");
            const shade = globalXmlParser.attr(node, "themeShade");
            return buildThemeColorReference(slot, tint, shade);
        }
    }
    class values {
        static themeValue(c, attr) {
            var val = globalXmlParser.attr(c, attr);
            return val ? `var(--docx-${val}-font)` : null;
        }
        static valueOfSize(c, attr) {
            var type = LengthUsage.Dxa;
            switch (globalXmlParser.attr(c, "type")) {
                case "dxa": break;
                case "pct":
                    type = LengthUsage.Percent;
                    break;
                case "auto": return "auto";
            }
            return globalXmlParser.lengthAttr(c, attr, type);
        }
        static valueOfMargin(c) {
            return globalXmlParser.lengthAttr(c, "w");
        }
        static applyShd(c, style) {
            const fill = xmlUtil.colorAttr(c, "fill", null, autos.shd, "themeFill");
            const color = xmlUtil.colorAttr(c, "color", null, autos.shd, "themeColor");
            const val = globalXmlParser.attr(c, "val");
            if (fill != null) {
                style["background-color"] = fill;
                const tref = xmlUtil.themeColorReference(c, "themeFill", "fill");
                if (tref)
                    style["$themeColor-background-color"] = tref;
            }
            if (!val || val === "clear" || val === "nil")
                return;
            const SHD_PATTERN_RE = /^(pct\d{1,2}|thin[A-Z][A-Za-z]+|[a-z][A-Za-z]+)$/;
            if (!SHD_PATTERN_RE.test(val))
                return;
            const base = fill ?? "transparent";
            const fg = color ?? "black";
            const pctMatch = /^pct(\d{1,2})$/.exec(val);
            if (pctMatch) {
                const pct = Math.min(100, Math.max(0, parseInt(pctMatch[1], 10)));
                style["background-color"] = `color-mix(in srgb, ${fg} ${pct}%, ${base})`;
                return;
            }
            const templates = {
                horzStripe: `repeating-linear-gradient(0deg, ${fg} 0 2px, ${base} 2px 6px)`,
                thinHorzStripe: `repeating-linear-gradient(0deg, ${fg} 0 1px, ${base} 1px 3px)`,
                vertStripe: `repeating-linear-gradient(90deg, ${fg} 0 2px, ${base} 2px 6px)`,
                thinVertStripe: `repeating-linear-gradient(90deg, ${fg} 0 1px, ${base} 1px 3px)`,
                diagStripe: `repeating-linear-gradient(45deg, ${fg} 0 2px, ${base} 2px 6px)`,
                thinDiagStripe: `repeating-linear-gradient(45deg, ${fg} 0 1px, ${base} 1px 3px)`,
                reverseDiagStripe: `repeating-linear-gradient(-45deg, ${fg} 0 2px, ${base} 2px 6px)`,
                thinReverseDiagStripe: `repeating-linear-gradient(-45deg, ${fg} 0 1px, ${base} 1px 3px)`,
            };
            const tpl = templates[val];
            if (tpl) {
                style["background-image"] = tpl;
                style["background-color"] = base;
                return;
            }
            const crossTemplates = {
                diagCross: `repeating-linear-gradient(45deg, ${fg} 0 2px, transparent 2px 6px), repeating-linear-gradient(-45deg, ${fg} 0 2px, transparent 2px 6px)`,
                thinDiagCross: `repeating-linear-gradient(45deg, ${fg} 0 1px, transparent 1px 3px), repeating-linear-gradient(-45deg, ${fg} 0 1px, transparent 1px 3px)`,
                horzCross: `repeating-linear-gradient(0deg, ${fg} 0 2px, transparent 2px 6px), repeating-linear-gradient(90deg, ${fg} 0 2px, transparent 2px 6px)`,
                thinHorzCross: `repeating-linear-gradient(0deg, ${fg} 0 1px, transparent 1px 3px), repeating-linear-gradient(90deg, ${fg} 0 1px, transparent 1px 3px)`,
            };
            const crossTpl = crossTemplates[val];
            if (crossTpl) {
                style["background-image"] = crossTpl;
                style["background-color"] = base;
            }
        }
        static valueOfBorder(c) {
            const rawVal = globalXmlParser.attr(c, "val");
            var type = values.parseBorderType(rawVal);
            if (type == "none")
                return "none";
            var color = xmlUtil.colorAttr(c, "color");
            var size = globalXmlParser.lengthAttr(c, "sz", LengthUsage.Border);
            const resolvedColor = (color == null || color == "auto") ? autos.borderColor : color;
            if (values.isArtBorder(rawVal)) {
                return `2pt solid ${resolvedColor}`;
            }
            return `${size} ${type} ${resolvedColor}`;
        }
        static isArtBorder(val) {
            return !!val && values.ART_BORDER_VALS.has(val);
        }
        static themeRefOfBorder(c) {
            if (values.parseBorderType(globalXmlParser.attr(c, "val")) == "none")
                return null;
            const literal = globalXmlParser.attr(c, "color");
            if (literal && literal !== "auto")
                return null;
            return xmlUtil.themeColorReference(c);
        }
        static parseBorderType(type) {
            switch (type) {
                case "single": return "solid";
                case "dashDotStroked": return "solid";
                case "dashed": return "dashed";
                case "dashSmallGap": return "dashed";
                case "dotDash": return "dotted";
                case "dotDotDash": return "dotted";
                case "dotted": return "dotted";
                case "double": return "double";
                case "doubleWave": return "double";
                case "inset": return "inset";
                case "nil": return "none";
                case "none": return "none";
                case "outset": return "outset";
                case "thick": return "solid";
                case "thickThinLargeGap": return "solid";
                case "thickThinMediumGap": return "solid";
                case "thickThinSmallGap": return "solid";
                case "thinThickLargeGap": return "solid";
                case "thinThickMediumGap": return "solid";
                case "thinThickSmallGap": return "solid";
                case "thinThickThinLargeGap": return "solid";
                case "thinThickThinMediumGap": return "solid";
                case "thinThickThinSmallGap": return "solid";
                case "threeDEmboss": return "solid";
                case "threeDEngrave": return "solid";
                case "triple": return "double";
                case "wave": return "solid";
            }
            return 'solid';
        }
        static valueOfTblLayout(c) {
            var type = globalXmlParser.attr(c, "val");
            return type == "fixed" ? "fixed" : "auto";
        }
        static classNameOfCnfStyle(c) {
            return classNameOfCnfStyle(c);
        }
        static valueOfJc(c) {
            var type = globalXmlParser.attr(c, "val");
            switch (type) {
                case "start":
                case "left": return "left";
                case "center": return "center";
                case "end":
                case "right": return "right";
                case "both": return "justify";
                case "distribute": return "justify";
            }
            return type;
        }
        static valueOfVertAlign(c, asTagName = false) {
            var type = globalXmlParser.attr(c, "val");
            switch (type) {
                case "subscript": return "sub";
                case "superscript": return asTagName ? "sup" : "super";
            }
            return asTagName ? null : type;
        }
        static valueOfTextAlignment(c) {
            var type = globalXmlParser.attr(c, "val");
            switch (type) {
                case "auto":
                case "baseline": return "baseline";
                case "top": return "top";
                case "center": return "middle";
                case "bottom": return "bottom";
            }
            return type;
        }
        static addSize(a, b) {
            if (a == null)
                return b;
            if (b == null)
                return a;
            return `calc(${a} + ${b})`;
        }
        static classNameOftblLook(c) {
            const val = globalXmlParser.hexAttr(c, "val", 0);
            let className = "";
            if (globalXmlParser.boolAttr(c, "firstRow") || (val & 0x0020))
                className += " first-row";
            if (globalXmlParser.boolAttr(c, "lastRow") || (val & 0x0040))
                className += " last-row";
            if (globalXmlParser.boolAttr(c, "firstColumn") || (val & 0x0080))
                className += " first-col";
            if (globalXmlParser.boolAttr(c, "lastColumn") || (val & 0x0100))
                className += " last-col";
            if (globalXmlParser.boolAttr(c, "noHBand") || (val & 0x0200))
                className += " no-hband";
            if (globalXmlParser.boolAttr(c, "noVBand") || (val & 0x0400))
                className += " no-vband";
            return className.trim();
        }
    }
    values.ART_BORDER_VALS = new Set([
        'apples', 'archedScallops', 'babyPants', 'babyRattle', 'balloons3Colors',
        'balloonsHotAir', 'basicBlackDashes', 'basicBlackDots', 'basicBlackSquares',
        'basicThinLines', 'basicWhiteDashes', 'basicWhiteDots', 'basicWhiteSquares',
        'basicWideInline', 'basicWideMidline', 'basicWideOutline', 'bats', 'birds',
        'birdsFlight', 'cabins', 'cakeSlice', 'candyCorn', 'celticKnotwork',
        'certificateBanner', 'chainLink', 'champagneBottle', 'checkedBarBlack',
        'checkedBarColor', 'checkered', 'christmasTree', 'circlesLines', 'circlesRectangles',
        'classicalWave', 'clocks', 'compass', 'confetti', 'confettiGrays', 'confettiOutline',
        'confettiStreamers', 'confettiWhite', 'cornerTriangles', 'couponCutoutDashes',
        'couponCutoutDots', 'crazyMaze', 'creaturesButterfly', 'creaturesFish',
        'creaturesInsects', 'creaturesLadyBug', 'crossStitch', 'cup', 'decoArch',
        'decoArchColor', 'decoBlocks', 'diamondsGray', 'doubleD', 'doubleDiamonds',
        'earth1', 'earth2', 'eclipsingSquares1', 'eclipsingSquares2', 'eggsBlack',
        'fans', 'film', 'firecrackers', 'flowersBlockPrint', 'flowersDaisies',
        'flowersModern1', 'flowersModern2', 'flowersPansy', 'flowersRedRose',
        'flowersRoses', 'flowersTeacup', 'flowersTiny', 'gems', 'gingerbreadMan',
        'gradient', 'handmade1', 'handmade2', 'heartBalloon', 'heartGray', 'hearts',
        'heebieJeebies', 'holly', 'houseFunky', 'hypnotic', 'iceCreamCones',
        'lightBulb', 'lightning1', 'lightning2', 'mapPins', 'mapleLeaf', 'mapleMuffins',
        'marquee', 'marqueeToothed', 'moons', 'mosaic', 'musicNotes', 'northwest',
        'ovals', 'packages', 'palmsBlack', 'palmsColor', 'paperClips', 'papyrus',
        'partyFavor', 'partyGlass', 'pencils', 'people', 'peopleHats', 'peopleWaving',
        'pinkFlowers', 'pumpkin1', 'pushPinNote1', 'pushPinNote2', 'pyramids',
        'pyramidsAbove', 'quadrants', 'rings', 'safari', 'sawtooth', 'sawtoothGray',
        'scaredCat', 'seattle', 'shadowedSquares', 'sharksTeeth', 'shorebirdTracks',
        'skyrocket', 'snowflakeFancy', 'snowflakes', 'sombrero', 'southwest', 'stars',
        'stars3d', 'starsBlack', 'starsShadowed', 'starsTop', 'sun', 'swirligig',
        'tornPaper', 'tornPaperBlack', 'trees', 'triangleParty', 'triangles',
        'tribal1', 'tribal2', 'tribal3', 'tribal4', 'tribal5', 'tribal6',
        'twistedLines1', 'twistedLines2', 'vine', 'waveline', 'weavingAngles',
        'weavingBraid', 'weavingRibbon', 'weavingStrips', 'whiteFlowers', 'woodwork',
        'xIllusions', 'zanyTriangles', 'zigZag', 'zigZagStitch'
    ]);

    const defaultTab = { pos: 0, leader: "none", style: "left" };
    const maxTabs = 50;
    function computePixelToPoint(container = document.body) {
        const temp = document.createElement("div");
        temp.style.width = '100pt';
        container.appendChild(temp);
        const result = 100 / temp.offsetWidth;
        container.removeChild(temp);
        return result;
    }
    function updateTabStop(elem, tabs, defaultTabSize, pixelToPoint = 72 / 96) {
        const p = elem.closest("p");
        const ebb = elem.getBoundingClientRect();
        const pbb = p.getBoundingClientRect();
        const pcs = getComputedStyle(p);
        const tabStops = tabs?.length > 0 ? tabs.map(t => ({
            pos: lengthToPoint(t.position),
            leader: t.leader,
            style: t.style
        })).sort((a, b) => a.pos - b.pos) : [defaultTab];
        const lastTab = tabStops[tabStops.length - 1];
        const pWidthPt = pbb.width * pixelToPoint;
        const size = lengthToPoint(defaultTabSize);
        let pos = lastTab.pos + size;
        if (pos < pWidthPt) {
            for (; pos < pWidthPt && tabStops.length < maxTabs; pos += size) {
                tabStops.push({ ...defaultTab, pos: pos });
            }
        }
        const marginLeft = parseFloat(pcs.marginLeft);
        const pOffset = pbb.left + marginLeft;
        const left = (ebb.left - pOffset) * pixelToPoint;
        const tab = tabStops.find(t => t.style != "clear" && t.pos > left);
        if (tab == null)
            return;
        let width = 1;
        if (tab.style == "right" || tab.style == "center") {
            const tabStops = Array.from(p.querySelectorAll(`.${elem.className}`));
            const nextIdx = tabStops.indexOf(elem) + 1;
            const range = document.createRange();
            range.setStart(elem, 1);
            if (nextIdx < tabStops.length) {
                range.setEndBefore(tabStops[nextIdx]);
            }
            else {
                range.setEndAfter(p);
            }
            const mul = tab.style == "center" ? 0.5 : 1;
            const nextBB = range.getBoundingClientRect();
            const offset = nextBB.left + mul * nextBB.width - (pbb.left - marginLeft);
            width = tab.pos - offset * pixelToPoint;
        }
        else {
            width = tab.pos - left;
        }
        elem.innerHTML = "&nbsp;";
        elem.style.textDecoration = "inherit";
        elem.style.wordSpacing = `${width.toFixed(0)}pt`;
        switch (tab.leader) {
            case "dot":
            case "middleDot":
                elem.style.textDecoration = "underline";
                elem.style.textDecorationStyle = "dotted";
                break;
            case "hyphen":
            case "heavy":
            case "underscore":
                elem.style.textDecoration = "underline";
                break;
        }
    }
    function lengthToPoint(length) {
        return parseFloat(length);
    }

    function parseFieldInstruction(raw) {
        const result = {
            code: '',
            switches: [],
            args: [],
            tokens: [],
            raw: raw ?? '',
        };
        if (!raw)
            return result;
        const tokens = [];
        let i = 0;
        const s = raw;
        while (i < s.length) {
            const ch = s[i];
            if (ch === ' ' || ch === '\t' || ch === '\r' || ch === '\n') {
                i++;
                continue;
            }
            if (ch === '"') {
                let buf = '';
                i++;
                while (i < s.length && s[i] !== '"') {
                    if (s[i] === '\\' && i + 1 < s.length && s[i + 1] === '"') {
                        buf += '"';
                        i += 2;
                    }
                    else {
                        buf += s[i++];
                    }
                }
                if (i < s.length)
                    i++;
                tokens.push(buf);
                continue;
            }
            let buf = '';
            while (i < s.length && !/\s/.test(s[i])) {
                buf += s[i++];
            }
            tokens.push(buf);
        }
        if (tokens.length === 0)
            return result;
        result.code = tokens[0].toUpperCase();
        for (let k = 1; k < tokens.length; k++) {
            const t = tokens[k];
            result.tokens.push(t);
            if (t.startsWith('\\') && t.length > 1)
                result.switches.push(t);
            else
                result.args.push(t);
        }
        return result;
    }

    var ns;
    (function (ns) {
        ns["html"] = "http://www.w3.org/1999/xhtml";
        ns["svg"] = "http://www.w3.org/2000/svg";
        ns["mathML"] = "http://www.w3.org/1998/Math/MathML";
    })(ns || (ns = {}));
    function h(elem) {
        if (isString(elem))
            return document.createTextNode(elem);
        if (elem instanceof Node)
            return elem;
        const { ns, tagName, className, style, children, ...props } = elem;
        if (tagName === "#fragment")
            return document.createDocumentFragment();
        if (tagName === "#comment")
            return document.createComment(children[0]);
        const result = (ns ? document.createElementNS(ns, tagName) : document.createElement(tagName));
        if (className)
            result.setAttribute("class", className);
        if (style) {
            if (isString(style)) {
                result.setAttribute("style", style);
            }
            else if (result.style) {
                Object.assign(result.style, style);
            }
        }
        if (props) {
            for (const [key, value] of Object.entries(props))
                if (value !== undefined)
                    result[key] = value;
        }
        if (children)
            children.forEach(c => result.appendChild(h(c)));
        return result;
    }
    function cx(...classNames) {
        return classNames.filter(Boolean).join(" ");
    }

    const VALID_COMMANDS = new Set(['M', 'L', 'C', 'Q', 'A', 'Z']);
    function scalePathD(d, pathW, pathH, renderW, renderH) {
        if (pathW <= 0 || pathH <= 0)
            return '';
        const sx = renderW / pathW;
        const sy = renderH / pathH;
        const tokens = d.split(/\s+/).filter(Boolean);
        let out = '';
        let cmd = '';
        let argIdx = 0;
        const argsPerCmd = {
            M: 2, L: 2, C: 6, Q: 4, A: 7, Z: 0,
        };
        for (const t of tokens) {
            if (VALID_COMMANDS.has(t)) {
                cmd = t;
                argIdx = 0;
                out += (out ? ' ' : '') + cmd;
                continue;
            }
            const n = Number(t);
            if (!Number.isFinite(n))
                continue;
            let v = n;
            if (cmd === 'A') {
                const ai = argIdx % 7;
                if (ai === 0)
                    v = n * sx;
                else if (ai === 1)
                    v = n * sy;
                else if (ai === 2)
                    v = n;
                else if (ai === 3)
                    v = n;
                else if (ai === 4)
                    v = n;
                else if (ai === 5)
                    v = n * sx;
                else
                    v = n * sy;
            }
            else if (argsPerCmd[cmd]) {
                const ai = argIdx % argsPerCmd[cmd];
                v = ai % 2 === 0 ? n * sx : n * sy;
            }
            out += ' ' + (Number.isFinite(v) ? v : 0);
            argIdx++;
        }
        return out;
    }
    function customGeometryToSvgPaths(custGeom, renderWidth, renderHeight) {
        if (!custGeom || !custGeom.paths || renderWidth <= 0 || renderHeight <= 0) {
            return [];
        }
        const out = [];
        for (const p of custGeom.paths) {
            if (!p || !p.d)
                continue;
            if (!Number.isFinite(p.w) || !Number.isFinite(p.h))
                continue;
            if (p.w <= 0 || p.h <= 0)
                continue;
            const scaled = scalePathD(p.d, p.w, p.h, renderWidth, renderHeight);
            if (scaled)
                out.push(scaled);
        }
        return out;
    }
    const DEFAULT_INSET_LR_EMU = 91440;
    const DEFAULT_INSET_TB_EMU = 45720;
    function presetGeometryToSvgPath(prst, w, h, adjustments) {
        if (!prst || w <= 0 || h <= 0)
            return null;
        const cx = w / 2;
        const cy = h / 2;
        const adj = (name, dflt) => {
            if (!adjustments)
                return dflt;
            const raw = adjustments[name];
            if (typeof raw !== 'number' || !Number.isFinite(raw))
                return dflt;
            return raw / 100000;
        };
        switch (prst) {
            case 'rect':
                return `M0,0 L${w},0 L${w},${h} L0,${h} Z`;
            case 'roundRect': {
                const frac = Math.max(0, Math.min(0.5, adj('adj', 0.1)));
                const r = Math.min(w, h) * frac;
                return (`M${r},0 L${w - r},0 Q${w},0 ${w},${r} ` +
                    `L${w},${h - r} Q${w},${h} ${w - r},${h} ` +
                    `L${r},${h} Q0,${h} 0,${h - r} ` +
                    `L0,${r} Q0,0 ${r},0 Z`);
            }
            case 'ellipse':
                return (`M0,${cy} A${cx},${cy} 0 1,0 ${w},${cy} ` +
                    `A${cx},${cy} 0 1,0 0,${cy} Z`);
            case 'triangle':
                return `M${cx},0 L${w},${h} L0,${h} Z`;
            case 'rtTriangle':
                return `M0,0 L0,${h} L${w},${h} Z`;
            case 'diamond':
                return `M${cx},0 L${w},${cy} L${cx},${h} L0,${cy} Z`;
            case 'parallelogram': {
                const skew = w * 0.25;
                return `M${skew},0 L${w},0 L${w - skew},${h} L0,${h} Z`;
            }
            case 'trapezoid': {
                const skew = w * 0.25;
                return `M${skew},0 L${w - skew},0 L${w},${h} L0,${h} Z`;
            }
            case 'pentagon': {
                const points = regularPolygon(5, cx, cy, w, h, -Math.PI / 2);
                return polygonToPath(points);
            }
            case 'hexagon': {
                const dx = w * 0.25;
                return (`M${dx},0 L${w - dx},0 L${w},${cy} ` +
                    `L${w - dx},${h} L${dx},${h} L0,${cy} Z`);
            }
            case 'octagon': {
                const d = w * 0.2929;
                const e = h * 0.2929;
                return (`M${d},0 L${w - d},0 L${w},${e} L${w},${h - e} ` +
                    `L${w - d},${h} L${d},${h} L0,${h - e} L0,${e} Z`);
            }
            case 'line':
                return `M0,0 L${w},${h}`;
            case 'rightArrow': {
                const head = w * 0.6;
                const shaftTop = h * 0.25;
                const shaftBot = h * 0.75;
                return (`M0,${shaftTop} L${head},${shaftTop} L${head},0 ` +
                    `L${w},${cy} L${head},${h} L${head},${shaftBot} ` +
                    `L0,${shaftBot} Z`);
            }
            case 'leftArrow': {
                const head = w * 0.4;
                const shaftTop = h * 0.25;
                const shaftBot = h * 0.75;
                return (`M${w},${shaftTop} L${head},${shaftTop} L${head},0 ` +
                    `L0,${cy} L${head},${h} L${head},${shaftBot} ` +
                    `L${w},${shaftBot} Z`);
            }
            case 'upArrow': {
                const head = h * 0.4;
                const shaftL = w * 0.25;
                const shaftR = w * 0.75;
                return (`M${shaftL},${h} L${shaftL},${head} L0,${head} ` +
                    `L${cx},0 L${w},${head} L${shaftR},${head} ` +
                    `L${shaftR},${h} Z`);
            }
            case 'downArrow': {
                const head = h * 0.6;
                const shaftL = w * 0.25;
                const shaftR = w * 0.75;
                return (`M${shaftL},0 L${shaftL},${head} L0,${head} ` +
                    `L${cx},${h} L${w},${head} L${shaftR},${head} ` +
                    `L${shaftR},0 Z`);
            }
            case 'leftRightArrow': {
                const headL = w * 0.2;
                const headR = w * 0.8;
                const shaftTop = h * 0.25;
                const shaftBot = h * 0.75;
                return (`M0,${cy} L${headL},0 L${headL},${shaftTop} ` +
                    `L${headR},${shaftTop} L${headR},0 L${w},${cy} ` +
                    `L${headR},${h} L${headR},${shaftBot} ` +
                    `L${headL},${shaftBot} L${headL},${h} Z`);
            }
            case 'wedgeRectCallout': {
                const tailX = cx + w * adj('adj1', -0.2);
                const tailY = cy + h * adj('adj2', 0.625);
                return (`M0,0 L${w},0 L${w},${h} ` +
                    `L${w * 0.5},${h} L${tailX},${tailY} L${w * 0.2},${h} ` +
                    `L0,${h} Z`);
            }
            case 'wedgeRoundRectCallout': {
                const r = Math.min(w, h) * Math.max(0, Math.min(0.5, adj('adj3', 0.1)));
                const tailX = cx + w * adj('adj1', -0.2);
                const tailY = cy + h * adj('adj2', 0.625);
                return (`M${r},0 L${w - r},0 Q${w},0 ${w},${r} ` +
                    `L${w},${h - r} Q${w},${h} ${w - r},${h} ` +
                    `L${w * 0.5},${h} L${tailX},${tailY} L${w * 0.2},${h} ` +
                    `L${r},${h} Q0,${h} 0,${h - r} ` +
                    `L0,${r} Q0,0 ${r},0 Z`);
            }
            case 'wedgeEllipseCallout': {
                const ang = (Math.PI * 5) / 4;
                const ex = cx + Math.cos(ang) * cx;
                const ey = cy + Math.sin(ang) * cy;
                const tailX = cx + w * adj('adj1', -0.2);
                const tailY = cy + h * adj('adj2', 0.625);
                return (`M${ex},${ey} L${tailX},${tailY} L${cx * 0.6},${h * 0.95} ` +
                    `A${cx},${cy} 0 1,1 ${ex},${ey} Z`);
            }
            case 'star5': {
                const ratio = Math.max(0.1, Math.min(0.9, adj('adj', 0.4)));
                const points = star(5, cx, cy, w, h, ratio, -Math.PI / 2);
                return polygonToPath(points);
            }
            case 'star6': {
                const ratio = Math.max(0.1, Math.min(0.9, adj('adj', 0.5)));
                const points = star(6, cx, cy, w, h, ratio, -Math.PI / 2);
                return polygonToPath(points);
            }
            case 'star8': {
                const ratio = Math.max(0.1, Math.min(0.9, adj('adj', 0.55)));
                const points = star(8, cx, cy, w, h, ratio, -Math.PI / 2);
                return polygonToPath(points);
            }
            case 'cloudCallout': {
                const tail1X = w * 0.1;
                const tail1Y = h * 1.05;
                const tail2X = w * -0.05;
                const tail2Y = h * 1.2;
                const tail1R = Math.min(w, h) * 0.05;
                const tail2R = Math.min(w, h) * 0.03;
                return (`M0,${cy} A${cx},${cy} 0 1,0 ${w},${cy} ` +
                    `A${cx},${cy} 0 1,0 0,${cy} Z ` +
                    `M${tail1X - tail1R},${tail1Y} ` +
                    `A${tail1R},${tail1R} 0 1,0 ${tail1X + tail1R},${tail1Y} ` +
                    `A${tail1R},${tail1R} 0 1,0 ${tail1X - tail1R},${tail1Y} Z ` +
                    `M${tail2X - tail2R},${tail2Y} ` +
                    `A${tail2R},${tail2R} 0 1,0 ${tail2X + tail2R},${tail2Y} ` +
                    `A${tail2R},${tail2R} 0 1,0 ${tail2X - tail2R},${tail2Y} Z`);
            }
            default:
                return null;
        }
    }
    function regularPolygon(n, cx, cy, w, h, startAng) {
        const rx = w / 2;
        const ry = h / 2;
        const pts = [];
        for (let i = 0; i < n; i++) {
            const ang = startAng + (i * 2 * Math.PI) / n;
            pts.push([cx + Math.cos(ang) * rx, cy + Math.sin(ang) * ry]);
        }
        return pts;
    }
    function star(points, cx, cy, w, h, innerRatio, startAng) {
        const rx = w / 2;
        const ry = h / 2;
        const irx = rx * innerRatio;
        const iry = ry * innerRatio;
        const total = points * 2;
        const pts = [];
        for (let i = 0; i < total; i++) {
            const outer = i % 2 === 0;
            const ang = startAng + (i * Math.PI) / points;
            const tx = outer ? rx : irx;
            const ty = outer ? ry : iry;
            pts.push([cx + Math.cos(ang) * tx, cy + Math.sin(ang) * ty]);
        }
        return pts;
    }
    function polygonToPath(points) {
        if (points.length === 0)
            return '';
        const [fx, fy] = points[0];
        let d = `M${fx},${fy}`;
        for (let i = 1; i < points.length; i++) {
            d += ` L${points[i][0]},${points[i][1]}`;
        }
        return d + ' Z';
    }
    function appendGradient(defs, grad, context) {
        if (!grad || !grad.stops || grad.stops.length === 0)
            return null;
        const tag = grad.kind === 'radial' ? 'radialGradient' : 'linearGradient';
        const el = document.createElementNS(ns.svg, tag);
        const id = safeId(context.nextId('docx-grad'));
        if (!id)
            return null;
        el.setAttribute('id', id);
        if (grad.kind === 'linear') {
            const ang = typeof grad.angle === 'number' && Number.isFinite(grad.angle)
                ? grad.angle
                : 0;
            const rad = (ang * Math.PI) / 180;
            const cx = 0.5, cy = 0.5;
            const x1 = cx - Math.cos(rad) * 0.5;
            const y1 = cy - Math.sin(rad) * 0.5;
            const x2 = cx + Math.cos(rad) * 0.5;
            const y2 = cy + Math.sin(rad) * 0.5;
            el.setAttribute('x1', x1.toFixed(4));
            el.setAttribute('y1', y1.toFixed(4));
            el.setAttribute('x2', x2.toFixed(4));
            el.setAttribute('y2', y2.toFixed(4));
        }
        else {
            el.setAttribute('cx', '0.5');
            el.setAttribute('cy', '0.5');
            el.setAttribute('r', '0.5');
        }
        for (const s of grad.stops) {
            const stop = document.createElementNS(ns.svg, 'stop');
            const pos = Math.max(0, Math.min(1, typeof s.pos === 'number' ? s.pos : 0));
            stop.setAttribute('offset', `${(pos * 100).toFixed(2)}%`);
            const hex = resolveColour(s.colour, context.themePalette);
            const c = sanitizeCssColor(hex);
            stop.setAttribute('stop-color', c ?? '#000000');
            const alpha = s.colour?.mods?.alpha;
            if (typeof alpha === 'number' && Number.isFinite(alpha)) {
                const o = Math.max(0, Math.min(1, alpha / 100000));
                stop.setAttribute('stop-opacity', o.toFixed(3));
            }
            el.appendChild(stop);
        }
        defs.appendChild(el);
        return id;
    }
    function appendPattern(defs, patt, context) {
        if (!patt || !patt.preset)
            return null;
        const fg = sanitizeCssColor(resolveColour(patt.fg, context.themePalette)) ?? '#000000';
        const bg = sanitizeCssColor(resolveColour(patt.bg, context.themePalette)) ?? '#FFFFFF';
        const id = safeId(context.nextId('docx-patt'));
        if (!id)
            return null;
        const size = 8;
        const pattern = document.createElementNS(ns.svg, 'pattern');
        pattern.setAttribute('id', id);
        pattern.setAttribute('width', String(size));
        pattern.setAttribute('height', String(size));
        pattern.setAttribute('patternUnits', 'userSpaceOnUse');
        const bgRect = document.createElementNS(ns.svg, 'rect');
        bgRect.setAttribute('x', '0');
        bgRect.setAttribute('y', '0');
        bgRect.setAttribute('width', String(size));
        bgRect.setAttribute('height', String(size));
        bgRect.setAttribute('fill', bg);
        pattern.appendChild(bgRect);
        const isDk = patt.preset.startsWith('dk');
        const sw = isDk ? 1.5 : 0.75;
        const addLine = (x1, y1, x2, y2) => {
            const line = document.createElementNS(ns.svg, 'line');
            line.setAttribute('x1', String(x1));
            line.setAttribute('y1', String(y1));
            line.setAttribute('x2', String(x2));
            line.setAttribute('y2', String(y2));
            line.setAttribute('stroke', fg);
            line.setAttribute('stroke-width', String(sw));
            pattern.appendChild(line);
        };
        switch (patt.preset) {
            case 'dkDnDiag':
            case 'ltDnDiag':
                addLine(-1, 0, size + 1, size + 2);
                addLine(-1, -size, size + 1, 2);
                addLine(-1, size, size + 1, size * 2 + 2);
                break;
            case 'dkUpDiag':
            case 'ltUpDiag':
                addLine(-1, size, size + 1, -2);
                addLine(-1, size * 2, size + 1, size - 2);
                addLine(-1, 0, size + 1, -size - 2);
                break;
            case 'dkHorz':
            case 'ltHorz':
                addLine(0, size / 2, size, size / 2);
                break;
            case 'dkVert':
            case 'ltVert':
                addLine(size / 2, 0, size / 2, size);
                break;
            case 'cross':
                addLine(0, size / 2, size, size / 2);
                addLine(size / 2, 0, size / 2, size);
                break;
            case 'diagCross':
                addLine(-1, -1, size + 1, size + 1);
                addLine(-1, size + 1, size + 1, -1);
                break;
            default:
                return null;
        }
        defs.appendChild(pattern);
        return id;
    }
    function appendEffects(defs, effects, widthPx, heightPx, context) {
        if (!effects)
            return null;
        const hasEffect = effects.outerShadow || effects.innerShadow || effects.softEdge || effects.glow;
        if (!hasEffect)
            return null;
        const id = safeId(context.nextId('docx-filter'));
        if (!id)
            return null;
        const filter = document.createElementNS(ns.svg, 'filter');
        filter.setAttribute('id', id);
        filter.setAttribute('x', '-50%');
        filter.setAttribute('y', '-50%');
        filter.setAttribute('width', '200%');
        filter.setAttribute('height', '200%');
        filter.setAttribute('filterUnits', 'objectBoundingBox');
        const layers = [];
        if (effects.outerShadow) {
            const s = effects.outerShadow;
            const { dx, dy } = shadowOffset(s.dir, s.dist);
            const blur = emuToShadowPx(s.blurRad) / 2;
            const colour = sanitizeCssColor(resolveColour(s.colour, context.themePalette)) ?? '#000000';
            const alpha = s.colour?.mods?.alpha;
            const opacity = typeof alpha === 'number' && Number.isFinite(alpha)
                ? Math.max(0, Math.min(1, alpha / 100000))
                : 0.5;
            const resultName = `outerShadow-${layers.length}`;
            const fe = document.createElementNS(ns.svg, 'feDropShadow');
            fe.setAttribute('in', 'SourceGraphic');
            fe.setAttribute('dx', dx.toFixed(2));
            fe.setAttribute('dy', dy.toFixed(2));
            fe.setAttribute('stdDeviation', blur.toFixed(2));
            fe.setAttribute('flood-color', colour);
            fe.setAttribute('flood-opacity', opacity.toFixed(3));
            fe.setAttribute('result', resultName);
            filter.appendChild(fe);
            layers.push(resultName);
        }
        if (effects.glow) {
            const rad = emuToShadowPx(effects.glow.rad);
            const colour = sanitizeCssColor(resolveColour(effects.glow.colour, context.themePalette)) ?? '#FFFF00';
            const alpha = effects.glow.colour?.mods?.alpha;
            const opacity = typeof alpha === 'number' && Number.isFinite(alpha)
                ? Math.max(0, Math.min(1, alpha / 100000))
                : 1;
            const blurResult = `glowBlur-${layers.length}`;
            const blur = document.createElementNS(ns.svg, 'feGaussianBlur');
            blur.setAttribute('in', 'SourceAlpha');
            blur.setAttribute('stdDeviation', rad.toFixed(2));
            blur.setAttribute('result', blurResult);
            filter.appendChild(blur);
            const floodResult = `glowFlood-${layers.length}`;
            const flood = document.createElementNS(ns.svg, 'feFlood');
            flood.setAttribute('flood-color', colour);
            flood.setAttribute('flood-opacity', opacity.toFixed(3));
            flood.setAttribute('result', floodResult);
            filter.appendChild(flood);
            const compResult = `glow-${layers.length}`;
            const comp = document.createElementNS(ns.svg, 'feComposite');
            comp.setAttribute('in', floodResult);
            comp.setAttribute('in2', blurResult);
            comp.setAttribute('operator', 'in');
            comp.setAttribute('result', compResult);
            filter.appendChild(comp);
            layers.push(compResult);
        }
        if (effects.innerShadow) {
            const s = effects.innerShadow;
            const { dx, dy } = shadowOffset(s.dir, s.dist);
            const blur = emuToShadowPx(s.blurRad) / 2;
            const colour = sanitizeCssColor(resolveColour(s.colour, context.themePalette)) ?? '#000000';
            const alpha = s.colour?.mods?.alpha;
            const opacity = typeof alpha === 'number' && Number.isFinite(alpha)
                ? Math.max(0, Math.min(1, alpha / 100000))
                : 0.6;
            const blurResult = `innerBlur-${layers.length}`;
            const blurEl = document.createElementNS(ns.svg, 'feGaussianBlur');
            blurEl.setAttribute('in', 'SourceAlpha');
            blurEl.setAttribute('stdDeviation', blur.toFixed(2));
            blurEl.setAttribute('result', blurResult);
            filter.appendChild(blurEl);
            const offsetResult = `innerOffset-${layers.length}`;
            const offsetEl = document.createElementNS(ns.svg, 'feOffset');
            offsetEl.setAttribute('in', blurResult);
            offsetEl.setAttribute('dx', dx.toFixed(2));
            offsetEl.setAttribute('dy', dy.toFixed(2));
            offsetEl.setAttribute('result', offsetResult);
            filter.appendChild(offsetEl);
            const invResult = `innerInvert-${layers.length}`;
            const invEl = document.createElementNS(ns.svg, 'feComposite');
            invEl.setAttribute('in', 'SourceAlpha');
            invEl.setAttribute('in2', offsetResult);
            invEl.setAttribute('operator', 'arithmetic');
            invEl.setAttribute('k2', '-1');
            invEl.setAttribute('k3', '1');
            invEl.setAttribute('result', invResult);
            filter.appendChild(invEl);
            const floodResult = `innerFlood-${layers.length}`;
            const flood = document.createElementNS(ns.svg, 'feFlood');
            flood.setAttribute('flood-color', colour);
            flood.setAttribute('flood-opacity', opacity.toFixed(3));
            flood.setAttribute('result', floodResult);
            filter.appendChild(flood);
            const clippedResult = `innerShadow-${layers.length}`;
            const clipEl = document.createElementNS(ns.svg, 'feComposite');
            clipEl.setAttribute('in', floodResult);
            clipEl.setAttribute('in2', invResult);
            clipEl.setAttribute('operator', 'in');
            clipEl.setAttribute('result', clippedResult);
            filter.appendChild(clipEl);
            const mergedResult = `innerShadowOver-${layers.length}`;
            const overEl = document.createElementNS(ns.svg, 'feComposite');
            overEl.setAttribute('in', clippedResult);
            overEl.setAttribute('in2', 'SourceGraphic');
            overEl.setAttribute('operator', 'over');
            overEl.setAttribute('result', mergedResult);
            filter.appendChild(overEl);
            layers.push(mergedResult);
        }
        if (effects.softEdge) {
            const rad = emuToShadowPx(effects.softEdge.rad);
            const resultName = `softEdge-${layers.length}`;
            const blur = document.createElementNS(ns.svg, 'feGaussianBlur');
            blur.setAttribute('in', 'SourceGraphic');
            blur.setAttribute('stdDeviation', rad.toFixed(2));
            blur.setAttribute('result', resultName);
            filter.appendChild(blur);
            layers.push(resultName);
        }
        const containsSource = !!effects.innerShadow || !!effects.softEdge;
        const merge = document.createElementNS(ns.svg, 'feMerge');
        for (const name of layers) {
            const node = document.createElementNS(ns.svg, 'feMergeNode');
            node.setAttribute('in', name);
            merge.appendChild(node);
        }
        if (!containsSource) {
            const src = document.createElementNS(ns.svg, 'feMergeNode');
            src.setAttribute('in', 'SourceGraphic');
            merge.appendChild(src);
        }
        filter.appendChild(merge);
        defs.appendChild(filter);
        return id;
    }
    function shadowOffset(dirDeg, distEmu) {
        const dir = typeof dirDeg === 'number' && Number.isFinite(dirDeg) ? dirDeg : 0;
        const distPx = emuToShadowPx(distEmu);
        const rad = (dir * Math.PI) / 180;
        return {
            dx: Math.cos(rad) * distPx,
            dy: Math.sin(rad) * distPx,
        };
    }
    function emuToShadowPx(emu) {
        if (typeof emu !== 'number' || !Number.isFinite(emu))
            return 0;
        return emu / 9525;
    }
    function safeId(id) {
        if (typeof id !== 'string')
            return null;
        if (!/^[a-z][a-z0-9-]*$/i.test(id))
            return null;
        return id;
    }
    let defaultIdCounter = 0;
    function defaultContext() {
        return {
            nextId(prefix) {
                defaultIdCounter += 1;
                return `${prefix}-${defaultIdCounter}`;
            },
        };
    }
    function renderShape(shape, emuToPx, renderText, ctx) {
        const context = ctx ?? defaultContext();
        const widthPx = emuToPx(shape.xfrm?.cx ?? 0);
        const heightPx = emuToPx(shape.xfrm?.cy ?? 0);
        const leftPx = emuToPx(shape.xfrm?.x ?? 0);
        const topPx = emuToPx(shape.xfrm?.y ?? 0);
        const wrapper = document.createElement('div');
        wrapper.className = 'docx-shape';
        wrapper.style.position = 'absolute';
        wrapper.style.left = `${leftPx.toFixed(2)}px`;
        wrapper.style.top = `${topPx.toFixed(2)}px`;
        wrapper.style.width = `${widthPx.toFixed(2)}px`;
        wrapper.style.height = `${heightPx.toFixed(2)}px`;
        const rot = shape.xfrm?.rot;
        if (rot && !Number.isNaN(rot)) {
            wrapper.style.transform = `rotate(${rot}deg)`;
        }
        const svg = document.createElementNS(ns.svg, 'svg');
        svg.setAttribute('xmlns', ns.svg);
        svg.setAttribute('viewBox', `0 0 ${widthPx} ${heightPx}`);
        svg.setAttribute('width', '100%');
        svg.setAttribute('height', '100%');
        svg.style.position = 'absolute';
        svg.style.inset = '0';
        svg.style.overflow = 'visible';
        const customDs = shape.custGeom && shape.custGeom.paths && shape.custGeom.paths.length > 0
            ? customGeometryToSvgPaths(shape.custGeom, widthPx, heightPx)
            : [];
        const presetD = customDs.length === 0
            ? (presetGeometryToSvgPath(shape.presetGeometry || 'rect', widthPx, heightPx, shape.presetAdjustments)
                ?? presetGeometryToSvgPath('rect', widthPx, heightPx))
            : null;
        const dStrings = customDs.length > 0 ? customDs : (presetD ? [presetD] : []);
        const defs = document.createElementNS(ns.svg, 'defs');
        let fillAttr = '#4472C4';
        if (shape.fill && shape.fill.type === 'solid') {
            const c = sanitizeCssColor(shape.fill.color);
            fillAttr = c ?? 'none';
        }
        else if (shape.fill && shape.fill.type === 'none') {
            fillAttr = 'none';
        }
        else if (shape.fill && shape.fill.type === 'gradient') {
            const id = appendGradient(defs, shape.fill.gradient, context);
            if (id)
                fillAttr = `url(#${id})`;
        }
        else if (shape.fill && shape.fill.type === 'pattern') {
            const id = appendPattern(defs, shape.fill.pattern, context);
            if (id)
                fillAttr = `url(#${id})`;
            else {
                const fg = resolveColour(shape.fill.pattern.fg, context.themePalette);
                const c = sanitizeCssColor(fg);
                if (c)
                    fillAttr = c;
            }
        }
        if (shape.presetGeometry === 'line' && customDs.length === 0) {
            fillAttr = 'none';
        }
        let strokeAttr = '#2F5496';
        let strokeWidthAttr = '1';
        if (shape.stroke) {
            strokeAttr = null;
            strokeWidthAttr = null;
            const stroke = sanitizeCssColor(shape.stroke.color);
            if (stroke)
                strokeAttr = stroke;
            if (shape.stroke.width != null && Number.isFinite(shape.stroke.width)) {
                const wPx = emuToPx(shape.stroke.width);
                strokeWidthAttr = `${wPx.toFixed(2)}`;
            }
        }
        const filterId = shape.effects
            ? appendEffects(defs, shape.effects, widthPx, heightPx, context)
            : null;
        if (defs.childNodes.length > 0)
            svg.appendChild(defs);
        for (const d of dStrings) {
            const path = document.createElementNS(ns.svg, 'path');
            path.setAttribute('d', d);
            path.setAttribute('fill', fillAttr);
            if (strokeAttr)
                path.setAttribute('stroke', strokeAttr);
            if (strokeWidthAttr)
                path.setAttribute('stroke-width', strokeWidthAttr);
            if (filterId)
                path.setAttribute('filter', `url(#${filterId})`);
            svg.appendChild(path);
        }
        if (shape.effects?.reflection) {
            const ref = shape.effects.reflection;
            const distPx = emuToShadowPx(ref.dist);
            const stA = typeof ref.stA === 'number' && Number.isFinite(ref.stA) ? ref.stA : 50000;
            const endA = typeof ref.endA === 'number' && Number.isFinite(ref.endA) ? ref.endA : 300;
            const startOpacity = Math.max(0, Math.min(1, stA / 100000)) * 0.5;
            const endOpacity = Math.max(0, Math.min(1, endA / 100000)) * 0.5;
            const dirDeg = typeof ref.dir === 'number' && Number.isFinite(ref.dir) ? ref.dir : 90;
            const dirRad = (dirDeg * Math.PI) / 180;
            const offsetX = distPx * Math.cos(dirRad);
            const offsetY = distPx * Math.sin(dirRad);
            const fadeDirDeg = typeof ref.fadeDir === 'number' && Number.isFinite(ref.fadeDir) ? ref.fadeDir : 90;
            const cssFadeDeg = ((fadeDirDeg + 90) % 360 + 360) % 360;
            const stPosRaw = typeof ref.stPos === 'number' && Number.isFinite(ref.stPos) ? ref.stPos : 0;
            const endPosRaw = typeof ref.endPos === 'number' && Number.isFinite(ref.endPos) ? ref.endPos : 100000;
            const stPosPct = Math.max(0, Math.min(100000, stPosRaw)) / 1000;
            const endPosPct = Math.max(0, Math.min(100000, endPosRaw)) / 1000;
            const reflectionSvg = svg.cloneNode(true);
            reflectionSvg.style.position = 'absolute';
            reflectionSvg.style.left = '0';
            reflectionSvg.style.top = '0';
            reflectionSvg.style.width = '100%';
            reflectionSvg.style.height = `${heightPx.toFixed(2)}px`;
            const shapeRotDeg = typeof rot === 'number' && Number.isFinite(rot) ? rot : 0;
            const rotWithShape = ref.rotWithShape !== false;
            const counterRot = (!rotWithShape && shapeRotDeg !== 0)
                ? `rotate(${(-shapeRotDeg).toFixed(3)}deg) `
                : '';
            reflectionSvg.style.transform =
                `${counterRot}scaleY(-1) ` +
                    `translate(${offsetX.toFixed(3)}px, ${(-heightPx - offsetY).toFixed(3)}px)`;
            reflectionSvg.style.transformOrigin = 'top left';
            reflectionSvg.style.pointerEvents = 'none';
            const maskGradient = `linear-gradient(${cssFadeDeg.toFixed(3)}deg, ` +
                `rgba(0,0,0,${startOpacity.toFixed(3)}) ${stPosPct.toFixed(3)}%, ` +
                `rgba(0,0,0,${endOpacity.toFixed(3)}) ${endPosPct.toFixed(3)}%)`;
            reflectionSvg.style.webkitMaskImage = maskGradient;
            reflectionSvg.style.maskImage = maskGradient;
            wrapper.appendChild(reflectionSvg);
        }
        wrapper.appendChild(svg);
        if (shape.txbxParagraphs && shape.txbxParagraphs.length > 0 && renderText) {
            const text = document.createElement('div');
            text.className = 'docx-shape-text';
            text.style.position = 'absolute';
            text.style.inset = '0';
            text.style.boxSizing = 'border-box';
            const lIns = shape.bodyPr?.lIns ?? DEFAULT_INSET_LR_EMU;
            const tIns = shape.bodyPr?.tIns ?? DEFAULT_INSET_TB_EMU;
            const rIns = shape.bodyPr?.rIns ?? DEFAULT_INSET_LR_EMU;
            const bIns = shape.bodyPr?.bIns ?? DEFAULT_INSET_TB_EMU;
            text.style.paddingLeft = `${emuToPx(lIns).toFixed(2)}px`;
            text.style.paddingTop = `${emuToPx(tIns).toFixed(2)}px`;
            text.style.paddingRight = `${emuToPx(rIns).toFixed(2)}px`;
            text.style.paddingBottom = `${emuToPx(bIns).toFixed(2)}px`;
            text.style.overflow = 'hidden';
            const rendered = renderText(shape.txbxParagraphs);
            for (const n of rendered) {
                if (n)
                    text.appendChild(n);
            }
            wrapper.appendChild(text);
        }
        return wrapper;
    }
    function renderShapeGroup(group, emuToPx, renderChild, _ctx) {
        const widthPx = emuToPx(group.xfrm?.cx ?? 0);
        const heightPx = emuToPx(group.xfrm?.cy ?? 0);
        const leftPx = emuToPx(group.xfrm?.x ?? 0);
        const topPx = emuToPx(group.xfrm?.y ?? 0);
        const wrapper = document.createElement('div');
        wrapper.className = 'docx-shape-group';
        wrapper.style.position = 'absolute';
        wrapper.style.left = `${leftPx.toFixed(2)}px`;
        wrapper.style.top = `${topPx.toFixed(2)}px`;
        wrapper.style.width = `${widthPx.toFixed(2)}px`;
        wrapper.style.height = `${heightPx.toFixed(2)}px`;
        const chOff = group.childOffset ?? { x: 0, y: 0, cx: group.xfrm?.cx ?? 0, cy: group.xfrm?.cy ?? 0 };
        const svg = document.createElementNS(ns.svg, 'svg');
        svg.setAttribute('xmlns', ns.svg);
        svg.setAttribute('viewBox', `${chOff.x} ${chOff.y} ${chOff.cx} ${chOff.cy}`);
        svg.setAttribute('width', '100%');
        svg.setAttribute('height', '100%');
        svg.style.overflow = 'visible';
        for (const child of group.children ?? []) {
            const node = renderChild(child);
            if (!node)
                continue;
            const x = child.xfrm?.x ?? 0;
            const y = child.xfrm?.y ?? 0;
            const cx = child.xfrm?.cx ?? chOff.cx;
            const cy = child.xfrm?.cy ?? chOff.cy;
            const fo = document.createElementNS(ns.svg, 'foreignObject');
            fo.setAttribute('x', String(x));
            fo.setAttribute('y', String(y));
            fo.setAttribute('width', String(cx));
            fo.setAttribute('height', String(cy));
            if (node instanceof HTMLElement) {
                node.style.position = 'relative';
                node.style.left = '0';
                node.style.top = '0';
                node.style.width = '100%';
                node.style.height = '100%';
            }
            fo.appendChild(node);
            svg.appendChild(fo);
        }
        wrapper.appendChild(svg);
        return wrapper;
    }

    const SVG_NS = "http://www.w3.org/2000/svg";
    const VIEW_W = 600;
    const VIEW_H = 400;
    const PADDING = 16;
    const TITLE_HEIGHT = 28;
    const LEGEND_ROW_HEIGHT = 18;
    const AXIS_LABEL_WIDTH = 40;
    const AXIS_LABEL_HEIGHT = 24;
    const FONT_SIZE = 11;
    const DEFAULT_PALETTE = [
        "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000",
        "#5B9BD5", "#70AD47", "#264478", "#9E480E",
        "#636363", "#997300",
    ];
    const DEFAULT_AXIS_LINE = "#888";
    const DEFAULT_TICK_LABEL = "#555";
    function resolveAxisRef(ref, palette) {
        if (!ref)
            return null;
        if (ref.kind === "literal") {
            return sanitizeCssColor(ref.color);
        }
        const resolved = resolveSchemeColor(ref.slot, palette ?? undefined);
        return resolved ? sanitizeCssColor(resolved) : null;
    }
    function resolveAxisStyle(style, palette) {
        const s = style ?? { line: null, tickLabel: null, gridline: null };
        return {
            line: resolveAxisRef(s.line, palette) ?? DEFAULT_AXIS_LINE,
            tickLabel: resolveAxisRef(s.tickLabel, palette) ?? DEFAULT_TICK_LABEL,
            gridline: resolveAxisRef(s.gridline, palette),
        };
    }
    function renderChart(model, options) {
        const doc = document;
        const svg = doc.createElementNS(SVG_NS, "svg");
        svg.setAttribute("viewBox", `0 0 ${VIEW_W} ${VIEW_H}`);
        svg.setAttribute("role", "img");
        svg.setAttribute("preserveAspectRatio", "xMidYMid meet");
        svg.style.maxWidth = "100%";
        svg.style.height = "auto";
        svg.style.display = "block";
        const palette = options?.themePalette ?? null;
        const catAxis = resolveAxisStyle(model.catAxis, palette);
        const valAxis = resolveAxisStyle(model.valAxis, palette);
        const title = (model.title ?? "").trim();
        const titleBottom = title ? PADDING + TITLE_HEIGHT : PADDING;
        if (title) {
            const t = mkText(doc, VIEW_W / 2, PADDING + TITLE_HEIGHT - 10, title, {
                fontSize: 14, fontWeight: "600", anchor: "middle",
            });
            svg.appendChild(t);
        }
        const legendEntries = model.showLegend
            ? model.series.filter((s) => s.values.length > 0)
            : [];
        const legendLayout = layoutLegend(legendEntries, VIEW_W - 2 * PADDING, doc);
        const legendTop = VIEW_H - PADDING - legendLayout.height;
        const plotTop = titleBottom + 4;
        const plotBottom = legendLayout.height > 0
            ? legendTop - 4
            : VIEW_H - PADDING;
        switch (model.kind) {
            case "column":
            case "bar":
                renderBarChart(svg, doc, model, {
                    left: PADDING + AXIS_LABEL_WIDTH,
                    right: VIEW_W - PADDING,
                    top: plotTop,
                    bottom: plotBottom - AXIS_LABEL_HEIGHT,
                    horizontal: model.kind === "bar",
                }, { catAxis, valAxis });
                break;
            case "line":
                renderLineChart(svg, doc, model, {
                    left: PADDING + AXIS_LABEL_WIDTH,
                    right: VIEW_W - PADDING,
                    top: plotTop,
                    bottom: plotBottom - AXIS_LABEL_HEIGHT,
                }, { catAxis, valAxis });
                break;
            case "pie":
                renderPieChart(svg, doc, model, {
                    left: PADDING,
                    right: VIEW_W - PADDING,
                    top: plotTop,
                    bottom: plotBottom,
                });
                break;
            default:
                const msg = mkText(doc, VIEW_W / 2, VIEW_H / 2, "[chart]", {
                    fontSize: 12, anchor: "middle", fill: "#888",
                });
                svg.appendChild(msg);
                break;
        }
        for (const node of legendLayout.nodes(legendTop))
            svg.appendChild(node);
        return svg;
    }
    function renderBarChart(svg, doc, model, rect, axisColors) {
        const series = model.series.filter((s) => s.values.length > 0);
        if (series.length === 0)
            return;
        const categories = maxLengthCategories(series);
        const catCount = categories.length;
        if (catCount === 0)
            return;
        const stacked = model.grouping === "stacked" || model.grouping === "percentStacked";
        const percentStacked = model.grouping === "percentStacked";
        let valMin = 0;
        let valMax = 0;
        if (stacked) {
            for (let i = 0; i < catCount; i++) {
                let sum = 0;
                for (const s of series) {
                    const v = finite(s.values[i]);
                    if (v == null)
                        continue;
                    sum += v;
                }
                if (percentStacked)
                    sum = sum === 0 ? 0 : 1;
                valMax = Math.max(valMax, sum);
                valMin = Math.min(valMin, sum);
            }
        }
        else {
            for (const s of series) {
                for (const v of s.values) {
                    const f = finite(v);
                    if (f == null)
                        continue;
                    valMax = Math.max(valMax, f);
                    valMin = Math.min(valMin, f);
                }
            }
        }
        if (valMax === valMin)
            valMax = valMin + 1;
        const { min: yMin, max: yMax, ticks } = niceScale(valMin, valMax, 5);
        const horizontal = rect.horizontal;
        const plotW = rect.right - rect.left;
        const plotH = rect.bottom - rect.top;
        const valLineColor = axisColors.valAxis.line;
        const catLineColor = axisColors.catAxis.line;
        const valTickColor = axisColors.valAxis.tickLabel;
        const catTickColor = axisColors.catAxis.tickLabel;
        const valGrid = axisColors.valAxis.gridline;
        svg.appendChild(mkLine(doc, rect.left, rect.top, rect.left, rect.bottom, valLineColor));
        svg.appendChild(mkLine(doc, rect.left, rect.bottom, rect.right, rect.bottom, catLineColor));
        for (const t of ticks) {
            if (horizontal) {
                const x = rect.left + ((t - yMin) / (yMax - yMin)) * plotW;
                if (valGrid && t !== yMin) {
                    svg.appendChild(mkLine(doc, x, rect.top, x, rect.bottom, valGrid));
                }
                svg.appendChild(mkLine(doc, x, rect.bottom, x, rect.bottom + 4, valLineColor));
                svg.appendChild(mkText(doc, x, rect.bottom + 16, formatTick(t), {
                    fontSize: FONT_SIZE, anchor: "middle", fill: valTickColor,
                }));
            }
            else {
                const y = rect.bottom - ((t - yMin) / (yMax - yMin)) * plotH;
                if (valGrid && t !== yMin) {
                    svg.appendChild(mkLine(doc, rect.left, y, rect.right, y, valGrid));
                }
                svg.appendChild(mkLine(doc, rect.left - 4, y, rect.left, y, valLineColor));
                svg.appendChild(mkText(doc, rect.left - 6, y + 4, formatTick(t), {
                    fontSize: FONT_SIZE, anchor: "end", fill: valTickColor,
                }));
            }
        }
        const slotSize = (horizontal ? plotH : plotW) / catCount;
        const groupPad = Math.min(slotSize * 0.15, 8);
        const innerGroup = slotSize - 2 * groupPad;
        const barsPerSlot = stacked ? 1 : series.length;
        const barSize = Math.max(1, innerGroup / Math.max(1, barsPerSlot));
        for (let i = 0; i < catCount; i++) {
            if (horizontal) {
                const yMid = rect.top + slotSize * (i + 0.5);
                svg.appendChild(mkText(doc, rect.left - 6, yMid + 4, categories[i] ?? "", {
                    fontSize: FONT_SIZE, anchor: "end", fill: catTickColor,
                }));
            }
            else {
                const xMid = rect.left + slotSize * (i + 0.5);
                svg.appendChild(mkText(doc, xMid, rect.bottom + 16, categories[i] ?? "", {
                    fontSize: FONT_SIZE, anchor: "middle", fill: catTickColor,
                }));
            }
            if (stacked) {
                let cumulative = 0;
                let slotTotal = 0;
                if (percentStacked) {
                    for (const s of series) {
                        const v = finite(s.values[i]);
                        if (v == null)
                            continue;
                        slotTotal += v;
                    }
                }
                for (let si = 0; si < series.length; si++) {
                    const raw = finite(series[si].values[i]);
                    if (raw == null || raw === 0)
                        continue;
                    const value = percentStacked
                        ? (slotTotal === 0 ? 0 : raw / slotTotal)
                        : raw;
                    const color = pointColor(series[si], i, si);
                    appendBar(svg, doc, {
                        horizontal, rect, yMin, yMax, plotW, plotH,
                        slotIndex: i, slotSize, barIndex: 0,
                        barSize: innerGroup, groupPad,
                        start: cumulative, end: cumulative + value,
                        color,
                    });
                    cumulative += value;
                }
            }
            else {
                for (let si = 0; si < series.length; si++) {
                    const value = finite(series[si].values[i]);
                    if (value == null)
                        continue;
                    const color = pointColor(series[si], i, si);
                    appendBar(svg, doc, {
                        horizontal, rect, yMin, yMax, plotW, plotH,
                        slotIndex: i, slotSize, barIndex: si,
                        barSize, groupPad,
                        start: 0, end: value,
                        color,
                    });
                }
            }
        }
    }
    function appendBar(svg, doc, p) {
        const scale = (v) => (v - p.yMin) / (p.yMax - p.yMin);
        const s0 = scale(Math.min(p.start, p.end));
        const s1 = scale(Math.max(p.start, p.end));
        if (p.horizontal) {
            const y = p.rect.top + p.slotSize * p.slotIndex + p.groupPad
                + p.barSize * p.barIndex;
            const x0 = p.rect.left + s0 * p.plotW;
            const x1 = p.rect.left + s1 * p.plotW;
            const rect = doc.createElementNS(SVG_NS, "rect");
            rect.setAttribute("x", fmt(Math.min(x0, x1)));
            rect.setAttribute("y", fmt(y));
            rect.setAttribute("width", fmt(Math.max(0, Math.abs(x1 - x0))));
            rect.setAttribute("height", fmt(Math.max(0, p.barSize)));
            rect.setAttribute("fill", p.color);
            svg.appendChild(rect);
        }
        else {
            const x = p.rect.left + p.slotSize * p.slotIndex + p.groupPad
                + p.barSize * p.barIndex;
            const y0 = p.rect.bottom - s0 * p.plotH;
            const y1 = p.rect.bottom - s1 * p.plotH;
            const rect = doc.createElementNS(SVG_NS, "rect");
            rect.setAttribute("x", fmt(x));
            rect.setAttribute("y", fmt(Math.min(y0, y1)));
            rect.setAttribute("width", fmt(Math.max(0, p.barSize)));
            rect.setAttribute("height", fmt(Math.max(0, Math.abs(y1 - y0))));
            rect.setAttribute("fill", p.color);
            svg.appendChild(rect);
        }
    }
    function renderLineChart(svg, doc, model, rect, axisColors) {
        const series = model.series.filter((s) => s.values.length > 0);
        if (series.length === 0)
            return;
        const categories = maxLengthCategories(series);
        const catCount = categories.length;
        if (catCount === 0)
            return;
        let valMin = Infinity;
        let valMax = -Infinity;
        for (const s of series) {
            for (const v of s.values) {
                const f = finite(v);
                if (f == null)
                    continue;
                if (f < valMin)
                    valMin = f;
                if (f > valMax)
                    valMax = f;
            }
        }
        if (!Number.isFinite(valMin) || !Number.isFinite(valMax))
            return;
        if (valMin > 0 && valMin / (valMax - valMin || 1) < 0.5)
            valMin = 0;
        if (valMax === valMin)
            valMax = valMin + 1;
        const { min: yMin, max: yMax, ticks } = niceScale(valMin, valMax, 5);
        const plotW = rect.right - rect.left;
        const plotH = rect.bottom - rect.top;
        const valLineColor = axisColors.valAxis.line;
        const catLineColor = axisColors.catAxis.line;
        const valTickColor = axisColors.valAxis.tickLabel;
        const catTickColor = axisColors.catAxis.tickLabel;
        const valGrid = axisColors.valAxis.gridline;
        svg.appendChild(mkLine(doc, rect.left, rect.top, rect.left, rect.bottom, valLineColor));
        svg.appendChild(mkLine(doc, rect.left, rect.bottom, rect.right, rect.bottom, catLineColor));
        for (const t of ticks) {
            const y = rect.bottom - ((t - yMin) / (yMax - yMin)) * plotH;
            if (valGrid && t !== yMin) {
                svg.appendChild(mkLine(doc, rect.left, y, rect.right, y, valGrid));
            }
            svg.appendChild(mkLine(doc, rect.left - 4, y, rect.left, y, valLineColor));
            svg.appendChild(mkText(doc, rect.left - 6, y + 4, formatTick(t), {
                fontSize: FONT_SIZE, anchor: "end", fill: valTickColor,
            }));
        }
        const xStep = catCount > 1 ? plotW / (catCount - 1) : 0;
        for (let i = 0; i < catCount; i++) {
            const x = rect.left + xStep * i;
            svg.appendChild(mkText(doc, x, rect.bottom + 16, categories[i] ?? "", {
                fontSize: FONT_SIZE, anchor: "middle", fill: catTickColor,
            }));
        }
        for (let si = 0; si < series.length; si++) {
            const s = series[si];
            const lineColor = seriesColor(s, si);
            const entries = [];
            for (let i = 0; i < catCount; i++) {
                const v = finite(s.values[i]);
                if (v == null)
                    continue;
                const x = catCount > 1 ? rect.left + xStep * i : rect.left + plotW / 2;
                const y = rect.bottom - ((v - yMin) / (yMax - yMin)) * plotH;
                entries.push({ x, y, pointIndex: i });
            }
            if (entries.length === 0)
                continue;
            const polyline = doc.createElementNS(SVG_NS, "polyline");
            polyline.setAttribute("points", entries.map((e) => `${fmt(e.x)},${fmt(e.y)}`).join(" "));
            polyline.setAttribute("fill", "none");
            polyline.setAttribute("stroke", lineColor);
            polyline.setAttribute("stroke-width", "2");
            svg.appendChild(polyline);
            for (const e of entries) {
                const circle = doc.createElementNS(SVG_NS, "circle");
                circle.setAttribute("cx", fmt(e.x));
                circle.setAttribute("cy", fmt(e.y));
                circle.setAttribute("r", "2.5");
                circle.setAttribute("fill", pointColor(s, e.pointIndex, si));
                svg.appendChild(circle);
            }
        }
    }
    function renderPieChart(svg, doc, model, rect) {
        const series = model.series.find((s) => s.values.length > 0);
        if (!series)
            return;
        const values = series.values.map((v) => finite(v) ?? 0);
        const total = values.reduce((a, b) => a + (b > 0 ? b : 0), 0);
        if (total <= 0)
            return;
        const cx = (rect.left + rect.right) / 2;
        const cy = (rect.top + rect.bottom) / 2;
        const r = Math.max(0, Math.min(rect.right - rect.left, rect.bottom - rect.top) / 2 - 8);
        if (r <= 0)
            return;
        let startAngle = -Math.PI / 2;
        for (let i = 0; i < values.length; i++) {
            const v = values[i];
            if (v <= 0)
                continue;
            const angle = (v / total) * 2 * Math.PI;
            const endAngle = startAngle + angle;
            const path = pieSlicePath(cx, cy, r, startAngle, endAngle);
            const slice = doc.createElementNS(SVG_NS, "path");
            slice.setAttribute("d", path);
            const override = series.dataPointOverrides?.get(i)?.color;
            const sanitisedOverride = override != null ? sanitizeCssColor(override) : null;
            const color = sanitisedOverride
                ?? sanitizeCssColor(series.color)
                ?? DEFAULT_PALETTE[i % DEFAULT_PALETTE.length];
            slice.setAttribute("fill", color);
            slice.setAttribute("stroke", "#fff");
            slice.setAttribute("stroke-width", "1");
            svg.appendChild(slice);
            startAngle = endAngle;
        }
    }
    function pieSlicePath(cx, cy, r, a0, a1) {
        if (a1 - a0 >= 2 * Math.PI - 1e-9) {
            return `M ${fmt(cx - r)} ${fmt(cy)} A ${fmt(r)} ${fmt(r)} 0 1 0 ${fmt(cx + r)} ${fmt(cy)} A ${fmt(r)} ${fmt(r)} 0 1 0 ${fmt(cx - r)} ${fmt(cy)} Z`;
        }
        const x0 = cx + r * Math.cos(a0);
        const y0 = cy + r * Math.sin(a0);
        const x1 = cx + r * Math.cos(a1);
        const y1 = cy + r * Math.sin(a1);
        const large = a1 - a0 > Math.PI ? 1 : 0;
        return `M ${fmt(cx)} ${fmt(cy)} L ${fmt(x0)} ${fmt(y0)} A ${fmt(r)} ${fmt(r)} 0 ${large} 1 ${fmt(x1)} ${fmt(y1)} Z`;
    }
    function layoutLegend(entries, maxWidth, doc) {
        if (entries.length === 0)
            return { height: 0, nodes: () => [] };
        const swatchW = 10;
        const gap = 6;
        const entryGap = 16;
        const estWidth = (title) => swatchW + gap + Math.max(12, title.length * FONT_SIZE * 0.55);
        const rows = [[]];
        let rowWidth = 0;
        for (let i = 0; i < entries.length; i++) {
            const e = entries[i];
            const w = estWidth(e.title || `Series ${i + 1}`);
            const needed = rowWidth === 0 ? w : w + entryGap;
            if (rowWidth + needed > maxWidth && rows[rows.length - 1].length > 0) {
                rows.push([]);
                rowWidth = 0;
            }
            rows[rows.length - 1].push({ entry: e, seriesIndex: i, width: w });
            rowWidth += needed;
        }
        const height = rows.length * LEGEND_ROW_HEIGHT;
        return {
            height,
            nodes: (top) => {
                const out = [];
                for (let r = 0; r < rows.length; r++) {
                    const row = rows[r];
                    const totalWidth = row.reduce((a, b, i) => a + b.width + (i === 0 ? 0 : entryGap), 0);
                    let x = (VIEW_W - totalWidth) / 2;
                    const y = top + r * LEGEND_ROW_HEIGHT + LEGEND_ROW_HEIGHT / 2;
                    const rowKey = `r${r}`;
                    for (const { entry, seriesIndex, width } of row) {
                        const color = seriesColor(entry, seriesIndex);
                        const swatch = doc.createElementNS(SVG_NS, "rect");
                        swatch.setAttribute("x", fmt(x));
                        swatch.setAttribute("y", fmt(y - swatchW / 2));
                        swatch.setAttribute("width", fmt(swatchW));
                        swatch.setAttribute("height", fmt(swatchW));
                        swatch.setAttribute("fill", color);
                        swatch.setAttribute("data-legend-row", rowKey);
                        out.push(swatch);
                        const label = mkText(doc, x + swatchW + gap, y + 4, entry.title || `Series ${seriesIndex + 1}`, {
                            fontSize: FONT_SIZE, anchor: "start", fill: "#333",
                        });
                        label.setAttribute("data-legend-entry", "1");
                        label.setAttribute("data-legend-est-w", fmt(width - swatchW - gap));
                        label.setAttribute("data-legend-row", rowKey);
                        out.push(label);
                        x += width + entryGap;
                    }
                }
                return out;
            },
        };
    }
    function seriesColor(s, index) {
        return sanitizeCssColor(s.color) ?? DEFAULT_PALETTE[index % DEFAULT_PALETTE.length];
    }
    function pointColor(s, pointIndex, seriesIndex) {
        const override = s.dataPointOverrides?.get(pointIndex)?.color;
        if (override != null) {
            const safe = sanitizeCssColor(override);
            if (safe)
                return safe;
        }
        return seriesColor(s, seriesIndex);
    }
    function finite(v) {
        return Number.isFinite(v) ? v : null;
    }
    function maxLengthCategories(series) {
        let longest = [];
        let count = 0;
        for (const s of series) {
            const len = Math.max(s.values.length, s.categories.length);
            if (len > count) {
                count = len;
                longest = s.categories.slice(0, len);
            }
        }
        while (longest.length < count)
            longest.push("");
        return longest;
    }
    function fmt(n) {
        if (!Number.isFinite(n))
            return "0";
        return Math.abs(n) < 0.005 ? "0" : n.toFixed(2);
    }
    function formatTick(n) {
        if (!Number.isFinite(n))
            return "";
        if (Math.abs(n) >= 1000)
            return n.toFixed(0);
        if (Math.abs(n) >= 10)
            return n.toFixed(1);
        return n.toFixed(2).replace(/\.?0+$/, "") || "0";
    }
    function niceScale(min, max, targetTicks) {
        const range = niceNum(max - min, false);
        const step = niceNum(range / Math.max(1, targetTicks - 1), true);
        const niceMin = Math.floor(min / step) * step;
        const niceMax = Math.ceil(max / step) * step;
        const ticks = [];
        const maxTicks = 32;
        for (let v = niceMin, i = 0; v <= niceMax + step * 0.5 && i < maxTicks; v += step, i++) {
            ticks.push(Number(v.toFixed(10)));
        }
        return { min: niceMin, max: niceMax, ticks };
    }
    function niceNum(range, round) {
        if (range <= 0)
            return 1;
        const exponent = Math.floor(Math.log10(range));
        const fraction = range / Math.pow(10, exponent);
        let niceFraction;
        if (round) {
            if (fraction < 1.5)
                niceFraction = 1;
            else if (fraction < 3)
                niceFraction = 2;
            else if (fraction < 7)
                niceFraction = 5;
            else
                niceFraction = 10;
        }
        else {
            if (fraction <= 1)
                niceFraction = 1;
            else if (fraction <= 2)
                niceFraction = 2;
            else if (fraction <= 5)
                niceFraction = 5;
            else
                niceFraction = 10;
        }
        return niceFraction * Math.pow(10, exponent);
    }
    function mkLine(doc, x1, y1, x2, y2, stroke = "#888") {
        const el = doc.createElementNS(SVG_NS, "line");
        el.setAttribute("x1", fmt(x1));
        el.setAttribute("y1", fmt(y1));
        el.setAttribute("x2", fmt(x2));
        el.setAttribute("y2", fmt(y2));
        el.setAttribute("stroke", stroke);
        el.setAttribute("stroke-width", "1");
        return el;
    }
    function mkText(doc, x, y, text, opts = {}) {
        const el = doc.createElementNS(SVG_NS, "text");
        el.setAttribute("x", fmt(x));
        el.setAttribute("y", fmt(y));
        if (opts.anchor)
            el.setAttribute("text-anchor", opts.anchor);
        el.setAttribute("font-size", String(opts.fontSize ?? FONT_SIZE));
        if (opts.fontWeight)
            el.setAttribute("font-weight", opts.fontWeight);
        if (opts.fill)
            el.setAttribute("fill", opts.fill);
        el.textContent = text;
        return el;
    }
    function scheduleLegendOverflowAdjust(svg) {
        if (typeof requestAnimationFrame !== "function")
            return;
        requestAnimationFrame(() => {
            if (!svg.isConnected)
                return;
            try {
                adjustLegendIfNeeded(svg);
            }
            catch {
            }
        });
    }
    function adjustLegendIfNeeded(svg) {
        const entries = Array.from(svg.querySelectorAll("text[data-legend-entry]"));
        if (entries.length === 0)
            return;
        let overflow = false;
        for (const entry of entries) {
            if (typeof entry.getBBox !== "function")
                return;
            const bbox = entry.getBBox();
            const estAttr = entry.getAttribute("data-legend-est-w");
            const est = estAttr ? parseFloat(estAttr) : NaN;
            if (!Number.isFinite(est))
                continue;
            if (bbox.width > est + 2) {
                overflow = true;
                break;
            }
        }
        if (!overflow)
            return;
        const byRow = new Map();
        const all = Array.from(svg.querySelectorAll("[data-legend-row]"));
        for (const node of all) {
            const key = node.getAttribute("data-legend-row") ?? "";
            const arr = byRow.get(key) ?? [];
            arr.push(node);
            byRow.set(key, arr);
        }
        for (const nodes of byRow.values()) {
            let rowRight = -Infinity;
            for (const node of nodes) {
                if (typeof node.getBBox !== "function")
                    continue;
                const bbox = node.getBBox();
                const right = bbox.x + bbox.width;
                if (right > rowRight)
                    rowRight = right;
            }
            if (!Number.isFinite(rowRight))
                continue;
            const allowed = VIEW_W - PADDING;
            if (rowRight <= allowed)
                continue;
            const shift = Math.min(rowRight - allowed, VIEW_W);
            for (const node of nodes) {
                const prev = node.getAttribute("transform") ?? "";
                const next = `${prev} translate(${(-shift).toFixed(2)},0)`.trim();
                node.setAttribute("transform", next);
            }
        }
    }
    const SUNBURST_LABEL_THRESHOLD = 0.12;
    const TREEMAP_LABEL_MIN_W = 32;
    const TREEMAP_LABEL_MIN_H = 14;
    function renderSunburst(model) {
        const doc = document;
        const svg = doc.createElementNS(SVG_NS, "svg");
        svg.setAttribute("viewBox", `0 0 ${VIEW_W} ${VIEW_H}`);
        svg.setAttribute("role", "img");
        svg.setAttribute("preserveAspectRatio", "xMidYMid meet");
        svg.style.maxWidth = "100%";
        svg.style.height = "auto";
        svg.style.display = "block";
        const title = (model.title ?? "").trim();
        const titleBottom = title ? PADDING + TITLE_HEIGHT : PADDING;
        if (title) {
            svg.appendChild(mkText(doc, VIEW_W / 2, PADDING + TITLE_HEIGHT - 10, title, {
                fontSize: 14, fontWeight: "600", anchor: "middle",
            }));
        }
        const plotTop = titleBottom + 4;
        const plotBottom = VIEW_H - PADDING;
        const cx = VIEW_W / 2;
        const cy = (plotTop + plotBottom) / 2;
        const outerR = Math.max(0, Math.min(VIEW_W - 2 * PADDING, plotBottom - plotTop) / 2 - 4);
        if (outerR <= 0 || model.root.value <= 0 || model.maxDepth === 0) {
            svg.appendChild(mkText(doc, VIEW_W / 2, VIEW_H / 2, title || "[sunburst]", {
                fontSize: 12, anchor: "middle", fill: "#888",
            }));
            return svg;
        }
        const innerR = Math.min(outerR * 0.15, 24);
        const ringThickness = (outerR - innerR) / model.maxDepth;
        renderSunburstNode(svg, doc, model.root, cx, cy, innerR, ringThickness, -Math.PI / 2, 2 * Math.PI, 0);
        return svg;
    }
    function renderSunburstNode(svg, doc, node, cx, cy, innerR, ringThickness, startAngle, sweep, paletteIdx) {
        const total = node.value > 0 ? node.value : sumChildValue(node);
        if (total <= 0 || node.children.length === 0)
            return;
        let angle = startAngle;
        for (let i = 0; i < node.children.length; i++) {
            const child = node.children[i];
            if (child.value <= 0)
                continue;
            const share = child.value / total;
            const slice = sweep * share;
            const r0 = innerR + ringThickness * child.level;
            const r1 = r0 + ringThickness;
            const basePaletteIdx = child.level === 0 ? i : paletteIdx;
            const color = sanitizeCssColor(child.color)
                ?? DEFAULT_PALETTE[basePaletteIdx % DEFAULT_PALETTE.length];
            const path = doc.createElementNS(SVG_NS, "path");
            path.setAttribute("d", sunburstArcPath(cx, cy, r0, r1, angle, angle + slice));
            path.setAttribute("fill", color);
            path.setAttribute("stroke", "#fff");
            path.setAttribute("stroke-width", "1");
            svg.appendChild(path);
            if (slice >= SUNBURST_LABEL_THRESHOLD && child.label) {
                const midAngle = angle + slice / 2;
                const midR = (r0 + r1) / 2;
                const tx = cx + midR * Math.cos(midAngle);
                const ty = cy + midR * Math.sin(midAngle);
                const text = mkText(doc, tx, ty + 4, child.label, {
                    fontSize: Math.min(FONT_SIZE, Math.max(8, ringThickness * 0.45)),
                    anchor: "middle",
                    fill: "#fff",
                });
                svg.appendChild(text);
            }
            renderSunburstNode(svg, doc, child, cx, cy, innerR, ringThickness, angle, slice, basePaletteIdx);
            angle += slice;
        }
    }
    function sumChildValue(node) {
        let total = 0;
        for (const c of node.children)
            total += c.value;
        return total;
    }
    function sunburstArcPath(cx, cy, r0, r1, a0, a1) {
        const large = a1 - a0 > Math.PI ? 1 : 0;
        const x0o = cx + r1 * Math.cos(a0);
        const y0o = cy + r1 * Math.sin(a0);
        const x1o = cx + r1 * Math.cos(a1);
        const y1o = cy + r1 * Math.sin(a1);
        const x0i = cx + r0 * Math.cos(a1);
        const y0i = cy + r0 * Math.sin(a1);
        const x1i = cx + r0 * Math.cos(a0);
        const y1i = cy + r0 * Math.sin(a0);
        return `M ${fmt(x0o)} ${fmt(y0o)}`
            + ` A ${fmt(r1)} ${fmt(r1)} 0 ${large} 1 ${fmt(x1o)} ${fmt(y1o)}`
            + ` L ${fmt(x0i)} ${fmt(y0i)}`
            + ` A ${fmt(r0)} ${fmt(r0)} 0 ${large} 0 ${fmt(x1i)} ${fmt(y1i)}`
            + ` Z`;
    }
    function renderTreemap(model) {
        const doc = document;
        const svg = doc.createElementNS(SVG_NS, "svg");
        svg.setAttribute("viewBox", `0 0 ${VIEW_W} ${VIEW_H}`);
        svg.setAttribute("role", "img");
        svg.setAttribute("preserveAspectRatio", "xMidYMid meet");
        svg.style.maxWidth = "100%";
        svg.style.height = "auto";
        svg.style.display = "block";
        const title = (model.title ?? "").trim();
        const titleBottom = title ? PADDING + TITLE_HEIGHT : PADDING;
        if (title) {
            svg.appendChild(mkText(doc, VIEW_W / 2, PADDING + TITLE_HEIGHT - 10, title, {
                fontSize: 14, fontWeight: "600", anchor: "middle",
            }));
        }
        const plotTop = titleBottom + 4;
        const plotBottom = VIEW_H - PADDING;
        const plot = {
            x: PADDING, y: plotTop,
            w: VIEW_W - 2 * PADDING, h: plotBottom - plotTop,
        };
        if (plot.h <= 0 || model.root.value <= 0) {
            svg.appendChild(mkText(doc, VIEW_W / 2, VIEW_H / 2, title || "[treemap]", {
                fontSize: 12, anchor: "middle", fill: "#888",
            }));
            return svg;
        }
        layoutSquarifiedTree(model.root, plot.x, plot.y, plot.w, plot.h);
        renderTreemapNodes(svg, doc, model.root, 0);
        return svg;
    }
    function layoutSquarifiedTree(node, x, y, w, h) {
        const ln = node;
        ln._x = x;
        ln._y = y;
        ln._w = w;
        ln._h = h;
        if (node.children.length === 0)
            return;
        if (w <= 0 || h <= 0) {
            for (const c of node.children)
                layoutSquarifiedTree(c, x, y, 0, 0);
            return;
        }
        const layout = squarifiedLayout(node.children, { x, y, width: w, height: h });
        for (const r of layout) {
            layoutSquarifiedTree(r.node, r.x, r.y, r.width, r.height);
        }
    }
    function squarifiedLayout(children, rect) {
        const out = [];
        if (children.length === 0)
            return out;
        if (!(rect.width > 0) || !(rect.height > 0))
            return out;
        if (children.length === 1) {
            out.push({
                node: children[0],
                x: rect.x, y: rect.y,
                width: rect.width, height: rect.height,
            });
            return out;
        }
        const items = children
            .map((node) => {
            const raw = parseFloat(node.value);
            const v = Number.isFinite(raw) && raw > 0 ? raw : 0;
            return { node, value: v };
        })
            .sort((a, b) => b.value - a.value);
        const total = items.reduce((s, i) => s + i.value, 0);
        if (total <= 0) {
            for (const it of items) {
                out.push({
                    node: it.node,
                    x: rect.x, y: rect.y,
                    width: 0, height: 0,
                });
            }
            return out;
        }
        const area = rect.width * rect.height;
        const scale = area / total;
        const scaled = items.map((i) => ({ node: i.node, area: i.value * scale }));
        squarifyInto(scaled, [], { ...rect }, out);
        return out;
    }
    function squarifyInto(remaining, row, rect, out) {
        while (true) {
            if (remaining.length === 0) {
                if (row.length > 0)
                    layoutRow(row, rect, out);
                return;
            }
            const w = Math.min(rect.width, rect.height);
            if (w <= 0) {
                for (const it of [...row, ...remaining]) {
                    out.push({ node: it.node, x: rect.x, y: rect.y, width: 0, height: 0 });
                }
                return;
            }
            const head = remaining[0];
            const extended = row.length === 0
                ? [head]
                : [...row, head];
            if (row.length === 0 || worstRatio(extended, w) <= worstRatio(row, w)) {
                row = extended;
                remaining = remaining.slice(1);
            }
            else {
                rect = layoutRow(row, rect, out);
                row = [];
            }
        }
    }
    function worstRatio(row, w) {
        let s = 0;
        let rmax = -Infinity;
        let rmin = Infinity;
        for (const r of row) {
            s += r.area;
            if (r.area > rmax)
                rmax = r.area;
            if (r.area < rmin)
                rmin = r.area;
        }
        if (s <= 0)
            return Infinity;
        const w2 = w * w;
        const s2 = s * s;
        const a = (w2 * rmax) / s2;
        const b = rmin > 0 ? s2 / (w2 * rmin) : Infinity;
        return Math.max(a, b);
    }
    function layoutRow(row, rect, out) {
        let sum = 0;
        for (const r of row)
            sum += r.area;
        if (sum <= 0) {
            for (const r of row) {
                out.push({ node: r.node, x: rect.x, y: rect.y, width: 0, height: 0 });
            }
            return rect;
        }
        const horizontal = rect.width >= rect.height;
        if (horizontal) {
            const stripW = sum / rect.height;
            let cy = rect.y;
            for (const r of row) {
                const hh = rect.height * (r.area / sum);
                out.push({
                    node: r.node,
                    x: rect.x, y: cy,
                    width: stripW, height: hh,
                });
                cy += hh;
            }
            return {
                x: rect.x + stripW, y: rect.y,
                width: Math.max(0, rect.width - stripW), height: rect.height,
            };
        }
        else {
            const stripH = sum / rect.width;
            let cx = rect.x;
            for (const r of row) {
                const ww = rect.width * (r.area / sum);
                out.push({
                    node: r.node,
                    x: cx, y: rect.y,
                    width: ww, height: stripH,
                });
                cx += ww;
            }
            return {
                x: rect.x, y: rect.y + stripH,
                width: rect.width, height: Math.max(0, rect.height - stripH),
            };
        }
    }
    function layoutTreemap(nodes, rect) {
        return squarifiedLayout(nodes, rect);
    }
    function renderTreemapNodes(svg, doc, node, paletteIdx) {
        for (let i = 0; i < node.children.length; i++) {
            const child = node.children[i];
            const seedIdx = child.level === 0 ? i : paletteIdx;
            if (child.children.length === 0) {
                renderTreemapLeaf(svg, doc, child, seedIdx);
            }
            else {
                renderTreemapNodes(svg, doc, child, seedIdx);
            }
        }
    }
    function renderTreemapLeaf(svg, doc, node, paletteIdx) {
        const ln = node;
        const x = ln._x ?? 0;
        const y = ln._y ?? 0;
        const w = ln._w ?? 0;
        const h = ln._h ?? 0;
        if (w <= 0 || h <= 0)
            return;
        const color = sanitizeCssColor(node.color)
            ?? DEFAULT_PALETTE[paletteIdx % DEFAULT_PALETTE.length];
        const rect = doc.createElementNS(SVG_NS, "rect");
        rect.setAttribute("x", fmt6(x));
        rect.setAttribute("y", fmt6(y));
        rect.setAttribute("width", fmt6(w));
        rect.setAttribute("height", fmt6(h));
        rect.setAttribute("fill", color);
        rect.setAttribute("stroke", "#fff");
        rect.setAttribute("stroke-width", "1");
        svg.appendChild(rect);
        if (node.label && w >= TREEMAP_LABEL_MIN_W && h >= TREEMAP_LABEL_MIN_H) {
            const label = mkText(doc, x + 4, y + 14, node.label, {
                fontSize: Math.min(FONT_SIZE, Math.max(9, Math.floor(h / 3))),
                anchor: "start",
                fill: "#fff",
            });
            svg.appendChild(label);
        }
    }
    function fmt6(n) {
        if (!Number.isFinite(n))
            return "0";
        return Number(n.toFixed(6)).toString();
    }
    const WATERFALL_POSITIVE = "#548235";
    const WATERFALL_NEGATIVE = "#C00000";
    const WATERFALL_TOTAL = "#4472C4";
    const HISTOGRAM_DEFAULT_BIN_COUNT = 10;
    const HISTOGRAM_MAX_BINS = 200;
    function chartExShell(title, emptyLabel) {
        const doc = document;
        const svg = doc.createElementNS(SVG_NS, "svg");
        svg.setAttribute("viewBox", `0 0 ${VIEW_W} ${VIEW_H}`);
        svg.setAttribute("role", "img");
        svg.setAttribute("preserveAspectRatio", "xMidYMid meet");
        svg.style.maxWidth = "100%";
        svg.style.height = "auto";
        svg.style.display = "block";
        const clean = (title ?? "").trim();
        const titleBottom = clean ? PADDING + TITLE_HEIGHT : PADDING;
        if (clean) {
            svg.appendChild(mkText(doc, VIEW_W / 2, PADDING + TITLE_HEIGHT - 10, clean, {
                fontSize: 14, fontWeight: "600", anchor: "middle",
            }));
        }
        const plotTop = titleBottom + 4;
        const plotBottom = VIEW_H - PADDING;
        const plot = {
            x: PADDING + AXIS_LABEL_WIDTH,
            y: plotTop,
            w: VIEW_W - 2 * PADDING - AXIS_LABEL_WIDTH,
            h: plotBottom - plotTop - AXIS_LABEL_HEIGHT,
        };
        return {
            svg, doc, plot,
            emptyIf(cond) {
                if (!cond)
                    return null;
                svg.appendChild(mkText(doc, VIEW_W / 2, VIEW_H / 2, clean || emptyLabel, {
                    fontSize: 12, anchor: "middle", fill: "#888",
                }));
                return svg;
            },
        };
    }
    function renderWaterfall(model) {
        const shell = chartExShell(model.title, "[waterfall]");
        const empty = shell.emptyIf(model.points.length === 0 || shell.plot.w <= 0 || shell.plot.h <= 0);
        if (empty)
            return empty;
        const { svg, doc, plot } = shell;
        const spans = [];
        let running = 0;
        for (const p of model.points) {
            let before;
            let after;
            if (p.type === "normal") {
                before = running;
                after = running + p.value;
                running = after;
            }
            else if (p.type === "subtotal") {
                before = 0;
                after = running + p.value;
                running = after;
            }
            else {
                before = 0;
                after = running + p.value;
            }
            spans.push({ before, after });
        }
        let valMin = 0;
        let valMax = 0;
        for (const s of spans) {
            if (s.before < valMin)
                valMin = s.before;
            if (s.after < valMin)
                valMin = s.after;
            if (s.before > valMax)
                valMax = s.before;
            if (s.after > valMax)
                valMax = s.after;
        }
        if (valMax === valMin)
            valMax = valMin + 1;
        const { min: yMin, max: yMax, ticks } = niceScale(valMin, valMax, 5);
        svg.appendChild(mkLine(doc, plot.x, plot.y, plot.x, plot.y + plot.h, DEFAULT_AXIS_LINE));
        svg.appendChild(mkLine(doc, plot.x, plot.y + plot.h, plot.x + plot.w, plot.y + plot.h, DEFAULT_AXIS_LINE));
        for (const t of ticks) {
            const y = plot.y + plot.h - ((t - yMin) / (yMax - yMin)) * plot.h;
            svg.appendChild(mkLine(doc, plot.x - 4, y, plot.x, y, DEFAULT_AXIS_LINE));
            svg.appendChild(mkText(doc, plot.x - 6, y + 4, formatTick(t), {
                fontSize: FONT_SIZE, anchor: "end", fill: DEFAULT_TICK_LABEL,
            }));
        }
        const n = model.points.length;
        const slotW = plot.w / n;
        const barPad = Math.min(slotW * 0.15, 8);
        const barW = Math.max(1, slotW - 2 * barPad);
        for (let i = 0; i < n; i++) {
            const p = model.points[i];
            const span = spans[i];
            const xMid = plot.x + slotW * (i + 0.5);
            svg.appendChild(mkText(doc, xMid, plot.y + plot.h + 16, p.label, {
                fontSize: FONT_SIZE, anchor: "middle", fill: DEFAULT_TICK_LABEL,
            }));
            const scaled = (v) => plot.y + plot.h
                - ((v - yMin) / (yMax - yMin)) * plot.h;
            const y0 = scaled(Math.max(span.before, span.after));
            const y1 = scaled(Math.min(span.before, span.after));
            const barH = Math.max(0, y1 - y0);
            const typeColor = p.type === "normal"
                ? (p.value >= 0 ? WATERFALL_POSITIVE : WATERFALL_NEGATIVE)
                : WATERFALL_TOTAL;
            const color = sanitizeCssColor(p.color) ?? typeColor;
            const rect = doc.createElementNS(SVG_NS, "rect");
            rect.setAttribute("x", fmt(xMid - barW / 2));
            rect.setAttribute("y", fmt(y0));
            rect.setAttribute("width", fmt(barW));
            rect.setAttribute("height", fmt(barH));
            rect.setAttribute("fill", color);
            svg.appendChild(rect);
        }
        return svg;
    }
    function renderFunnel(model) {
        const shell = chartExShell(model.title, "[funnel]");
        const n = model.points.length;
        let maxVal = 0;
        for (const p of model.points) {
            if (p.value > maxVal)
                maxVal = p.value;
        }
        const empty = shell.emptyIf(n === 0 || maxVal <= 0 || shell.plot.w <= 0 || shell.plot.h <= 0);
        if (empty)
            return empty;
        const { svg, doc, plot } = shell;
        const labelReserve = Math.min(160, plot.w * 0.33);
        const funnelW = Math.max(1, plot.w - labelReserve);
        const cx = plot.x + funnelW / 2;
        const bandH = plot.h / n;
        const vertPad = Math.min(bandH * 0.1, 6);
        const inner = bandH - 2 * vertPad;
        for (let i = 0; i < n; i++) {
            const p = model.points[i];
            const nextVal = i + 1 < n ? model.points[i + 1].value : p.value;
            const topHalf = (p.value / maxVal) * funnelW / 2;
            const botHalf = (Math.min(nextVal, p.value) / maxVal) * funnelW / 2;
            const y0 = plot.y + bandH * i + vertPad;
            const y1 = y0 + inner;
            const x0t = cx - topHalf;
            const x1t = cx + topHalf;
            const x0b = cx - botHalf;
            const x1b = cx + botHalf;
            const color = sanitizeCssColor(p.color)
                ?? DEFAULT_PALETTE[i % DEFAULT_PALETTE.length];
            const poly = doc.createElementNS(SVG_NS, "polygon");
            const points = [
                `${fmt(x0t)},${fmt(y0)}`,
                `${fmt(x1t)},${fmt(y0)}`,
                `${fmt(x1b)},${fmt(y1)}`,
                `${fmt(x0b)},${fmt(y1)}`,
            ].join(" ");
            poly.setAttribute("points", points);
            poly.setAttribute("fill", color);
            poly.setAttribute("stroke", "#fff");
            poly.setAttribute("stroke-width", "1");
            svg.appendChild(poly);
            const labelX = plot.x + funnelW + 8;
            const labelY = (y0 + y1) / 2 + 4;
            const labelText = p.label
                ? `${p.label}: ${formatTick(p.value)}`
                : formatTick(p.value);
            svg.appendChild(mkText(doc, labelX, labelY, labelText, {
                fontSize: FONT_SIZE, anchor: "start", fill: "#333",
            }));
        }
        return svg;
    }
    function renderHistogram(model) {
        const shell = chartExShell(model.title, "[histogram]");
        const values = model.values;
        const n = values.length;
        let minV = Infinity;
        let maxV = -Infinity;
        for (const v of values) {
            if (v < minV)
                minV = v;
            if (v > maxV)
                maxV = v;
        }
        const haveRange = n > 0 && Number.isFinite(minV) && Number.isFinite(maxV);
        const empty = shell.emptyIf(!haveRange || shell.plot.w <= 0 || shell.plot.h <= 0);
        if (empty)
            return empty;
        const { svg, doc, plot } = shell;
        const { underflow, overflow } = model.binning;
        const useUnderflow = underflow != null && underflow > minV;
        const useOverflow = overflow != null && overflow < maxV;
        const rangeLo = useUnderflow ? underflow : minV;
        let rangeHi = useOverflow ? overflow : maxV;
        if (rangeHi <= rangeLo)
            rangeHi = rangeLo + 1;
        let binSize = model.binning.binSize;
        let binCount = model.binning.binCount;
        if (binSize == null) {
            const count = binCount != null
                ? binCount
                : Math.max(1, Math.min(HISTOGRAM_MAX_BINS, HISTOGRAM_DEFAULT_BIN_COUNT));
            binSize = (rangeHi - rangeLo) / count;
        }
        if (!(binSize > 0))
            binSize = rangeHi - rangeLo;
        binCount = Math.max(1, Math.min(HISTOGRAM_MAX_BINS, Math.ceil((rangeHi - rangeLo) / binSize)));
        rangeHi = rangeLo + binSize * binCount;
        const bins = [];
        if (useUnderflow) {
            bins.push({
                lo: -Infinity, hi: rangeLo, count: 0,
                label: `<= ${formatTick(rangeLo)}`,
            });
        }
        for (let i = 0; i < binCount; i++) {
            const lo = rangeLo + binSize * i;
            const hi = lo + binSize;
            bins.push({
                lo, hi, count: 0,
                label: `${formatTick(lo)}-${formatTick(hi)}`,
            });
        }
        if (useOverflow) {
            bins.push({
                lo: rangeHi, hi: Infinity, count: 0,
                label: `> ${formatTick(rangeHi)}`,
            });
        }
        for (const v of values) {
            if (useUnderflow && v < rangeLo) {
                bins[0].count++;
                continue;
            }
            if (useOverflow && v > rangeHi) {
                bins[bins.length - 1].count++;
                continue;
            }
            let idx = Math.floor((v - rangeLo) / binSize);
            if (idx < 0)
                idx = 0;
            if (idx >= binCount)
                idx = binCount - 1;
            const offset = useUnderflow ? 1 : 0;
            bins[idx + offset].count++;
        }
        let maxCount = 0;
        for (const b of bins) {
            if (b.count > maxCount)
                maxCount = b.count;
        }
        if (maxCount === 0)
            maxCount = 1;
        const { min: yMin, max: yMax, ticks } = niceScale(0, maxCount, 5);
        svg.appendChild(mkLine(doc, plot.x, plot.y, plot.x, plot.y + plot.h, DEFAULT_AXIS_LINE));
        svg.appendChild(mkLine(doc, plot.x, plot.y + plot.h, plot.x + plot.w, plot.y + plot.h, DEFAULT_AXIS_LINE));
        for (const t of ticks) {
            const y = plot.y + plot.h - ((t - yMin) / (yMax - yMin)) * plot.h;
            svg.appendChild(mkLine(doc, plot.x - 4, y, plot.x, y, DEFAULT_AXIS_LINE));
            svg.appendChild(mkText(doc, plot.x - 6, y + 4, formatTick(t), {
                fontSize: FONT_SIZE, anchor: "end", fill: DEFAULT_TICK_LABEL,
            }));
        }
        const slotW = plot.w / bins.length;
        const barPad = Math.min(slotW * 0.1, 4);
        const barW = Math.max(1, slotW - 2 * barPad);
        const baseColor = sanitizeCssColor(model.seriesColor) ?? DEFAULT_PALETTE[0];
        for (let i = 0; i < bins.length; i++) {
            const b = bins[i];
            const xLeft = plot.x + slotW * i + barPad;
            const y0 = plot.y + plot.h - ((b.count - yMin) / (yMax - yMin)) * plot.h;
            const y1 = plot.y + plot.h;
            const override = model.dataPointOverrides.get(i);
            const color = (override ? sanitizeCssColor(override) : null) ?? baseColor;
            const rect = doc.createElementNS(SVG_NS, "rect");
            rect.setAttribute("x", fmt(xLeft));
            rect.setAttribute("y", fmt(y0));
            rect.setAttribute("width", fmt(barW));
            rect.setAttribute("height", fmt(Math.max(0, y1 - y0)));
            rect.setAttribute("fill", color);
            svg.appendChild(rect);
            svg.appendChild(mkText(doc, xLeft + barW / 2, plot.y + plot.h + 16, b.label, {
                fontSize: FONT_SIZE, anchor: "middle", fill: DEFAULT_TICK_LABEL,
            }));
        }
        return svg;
    }

    const OOX_CLASSES = {
        "wrapper": "oox-wrapper",
        "page": "oox-page",
        "paragraph": "oox-paragraph",
        "run": "oox-run",
        "heading": "oox-heading",
        "table": "oox-table",
        "table-row": "oox-table-row",
        "table-cell": "oox-table-cell",
        "image": "oox-image",
    };
    function addSharedClass(el, concept) {
        if (!el)
            return;
        el.classList.add(OOX_CLASSES[concept]);
    }

    const SAFE_HREF_SCHEMES = new Set(['http:', 'https:', 'mailto:', 'tel:', 'ftp:', 'ftps:']);
    function generateRenderSessionId() {
        try {
            const bytes = new Uint8Array(3);
            if (typeof globalThis.crypto?.getRandomValues === 'function') {
                globalThis.crypto.getRandomValues(bytes);
                return Array.from(bytes, b => b.toString(16).padStart(2, '0')).join('').slice(0, 5);
            }
        }
        catch { }
        return Math.floor(Math.random() * 0xfffff).toString(16).padStart(5, '0');
    }
    function isSafeHyperlinkHref(raw) {
        if (raw == null)
            return true;
        if (typeof raw !== 'string')
            return false;
        const trimmed = raw.trim();
        if (trimmed === '')
            return true;
        if (trimmed.startsWith('#'))
            return true;
        try {
            const parsed = new URL(trimmed, 'http://docxjs.invalid/');
            return SAFE_HREF_SCHEMES.has(parsed.protocol);
        }
        catch {
            return !/^[a-z][a-z0-9+.-]*:/i.test(trimmed);
        }
    }
    function isSafeCustomXmlXPath(xp) {
        if (typeof xp !== 'string')
            return false;
        if (xp.length === 0 || xp.length > 2000)
            return false;
        if (/\b(document|window|eval|Function|require|import)\b/i.test(xp))
            return false;
        if (/[;{}]/.test(xp))
            return false;
        return /^[\w\s\-./:[\]()=*|,!<>@"'$+]+$/.test(xp);
    }
    function safeEvaluateXPath(hostDoc, xmlDoc, xpath) {
        try {
            if (hostDoc?.evaluate) {
                const rootNs = xmlDoc?.documentElement?.namespaceURI ?? null;
                const resolver = rootNs ? (_) => rootNs : null;
                const FIRST_ORDERED_NODE_TYPE = 9;
                const res = hostDoc.evaluate(xpath, xmlDoc, resolver, FIRST_ORDERED_NODE_TYPE, null);
                const node = res?.singleNodeValue;
                if (node == null)
                    return '';
                return node.textContent ?? '';
            }
        }
        catch {
        }
        try {
            const steps = xpath.split('/').filter(s => s.length > 0);
            let current = xmlDoc?.documentElement;
            if (!current)
                return null;
            const first = steps[0]?.replace(/\[.*\]$/, '').split(':').pop();
            if (first && first === current.localName) {
                steps.shift();
            }
            for (const step of steps) {
                const localName = step.replace(/\[.*\]$/, '').split(':').pop();
                if (!localName)
                    return null;
                let next = null;
                for (let i = 0; i < current.childNodes.length; i++) {
                    const c = current.childNodes[i];
                    if (c.nodeType === 1 && c.localName === localName) {
                        next = c;
                        break;
                    }
                }
                if (!next)
                    return '';
                current = next;
            }
            return current.textContent ?? '';
        }
        catch {
            return null;
        }
    }
    function complexFieldCharType(elem) {
        if (!elem || elem.type !== DomType.Run)
            return null;
        const run = elem;
        if (!run.fieldRun || !run.children)
            return null;
        for (const c of run.children) {
            if (c.type === DomType.ComplexField)
                return c.charType;
        }
        return null;
    }
    function isComplexFieldBeginRun(elem) {
        return complexFieldCharType(elem) === 'begin';
    }
    const HYPERLINK_VALUE_SWITCHES = new Set(['\\o', '\\t']);
    function extractSwitchValue(parsed, switchName) {
        const tokens = parsed.tokens ?? [];
        const target = switchName.toLowerCase();
        for (let i = 0; i < tokens.length; i++) {
            if (tokens[i].toLowerCase() === target) {
                const next = tokens[i + 1];
                if (next && !next.startsWith('\\'))
                    return next;
                return null;
            }
        }
        return null;
    }
    function firstNonSwitchArg(parsed) {
        const tokens = parsed.tokens ?? [];
        let i = 0;
        while (i < tokens.length) {
            const t = tokens[i];
            if (t.startsWith('\\')) {
                if (HYPERLINK_VALUE_SWITCHES.has(t.toLowerCase())) {
                    const next = tokens[i + 1];
                    if (next && !next.startsWith('\\')) {
                        i += 2;
                        continue;
                    }
                }
                i += 1;
                continue;
            }
            return t;
        }
        return null;
    }
    function resolveNaryLimitTag(limLoc, opChar) {
        if (limLoc === "subSup")
            return "msubsup";
        if (limLoc === "undOvr")
            return "munderover";
        const BIG_OP = new Set(['∑', '∏', '⋃', '⋂', '⨁', '⨂', '⨀']);
        return BIG_OP.has(opChar) ? "munderover" : "msubsup";
    }
    function resolveGroupTag(pos, vertJc) {
        const hasTop = pos === "top" || vertJc === "top";
        const hasBot = pos === "bot" || pos === "bottom" || vertJc === "bot" || vertJc === "bottom";
        if (hasTop && hasBot)
            return "munderover";
        if (hasTop)
            return "mover";
        if (hasBot)
            return "munder";
        return "munderover";
    }
    const BCP47_RE = /^[A-Za-z]{1,8}(-[A-Za-z0-9]{1,8})*$/;
    function isValidBcp47LanguageTag(value) {
        return typeof value === 'string' && BCP47_RE.test(value);
    }
    function getHeadingTagName(paragraph, stylesMap) {
        if (!paragraph)
            return 'p';
        const level = resolveHeadingLevel(paragraph, stylesMap);
        if (level == null)
            return 'p';
        if (!Number.isInteger(level) || level < 0 || level > 5)
            return 'p';
        return `h${level + 1}`;
    }
    function resolveHeadingLevel(paragraph, stylesMap) {
        if (Number.isInteger(paragraph.outlineLevel) && paragraph.outlineLevel >= 0 && paragraph.outlineLevel <= 8) {
            return paragraph.outlineLevel;
        }
        const style = paragraph.styleName && stylesMap?.[paragraph.styleName];
        if (!style)
            return null;
        const styleLevel = style.paragraphProps?.outlineLevel;
        if (Number.isInteger(styleLevel) && styleLevel >= 0 && styleLevel <= 8) {
            return styleLevel;
        }
        const seen = new Set();
        let cursor = style;
        while (cursor && !seen.has(cursor.id)) {
            seen.add(cursor.id);
            const nameLevel = headingLevelFromName(cursor.name) ?? headingLevelFromName(cursor.id);
            if (nameLevel != null)
                return nameLevel;
            cursor = cursor.basedOn ? stylesMap?.[cursor.basedOn] : undefined;
        }
        return null;
    }
    function headingLevelFromName(name) {
        if (!name)
            return null;
        const m = /^heading\s*([1-9])$/i.exec(name.trim());
        if (!m)
            return null;
        const n = Number(m[1]);
        return n - 1;
    }
    class HtmlRenderer {
        constructor() {
            this.className = "docx";
            this.styleMap = {};
            this.fontAltNames = new Map();
            this.currentPart = null;
            this.tableVerticalMerges = [];
            this.currentVerticalMerge = null;
            this.tableCellPositions = [];
            this.currentCellPosition = null;
            this.tableBandSizes = [];
            this.currentTableBandSizes = { col: 1, row: 1 };
            this.currentRowIsHeader = false;
            this.currentSectProps = null;
            this._shapeIdCounter = 0;
            this.renderSessionId = '';
            this.footnoteMap = {};
            this.endnoteMap = {};
            this.currentEndnoteIds = [];
            this.footnoteRefCount = 0;
            this.endnoteRefCount = 0;
            this.usedHederFooterParts = [];
            this.currentTabs = [];
            this.lineNumberingArticleSeq = 0;
            this.commentMap = {};
            this.commentAnchorElements = {};
            this.sidebarContainer = null;
            this.sidebarCommentElements = {};
            this.revisionCardElements = new Map();
            this.changeAuthorIndex = new Map();
            this.changeElements = [];
            this.changeMeta = [];
            this.moveElements = new Map();
            this.tasks = [];
            this.postRenderTasks = [];
            this.h = h;
        }
        get useSidebar() {
            return this.options.renderComments && (this.options.comments?.sidebar !== false);
        }
        get useHighlight() {
            return this.options.renderComments && (this.options.comments?.highlight !== false);
        }
        get sidebarLayout() {
            return this.options.comments?.layout === 'packed' ? 'packed' : 'anchored';
        }
        get showChanges() {
            return !!this.options.changes?.show;
        }
        async render(document, options) {
            this.document = document;
            this.options = options;
            this.className = options.className;
            this.rootSelector = options.inWrapper ? `.${this.className}-wrapper` : ':root';
            this.h = options.h ?? h;
            this.renderSessionId = generateRenderSessionId();
            this.styleMap = null;
            this.tasks = [];
            this.commentAnchorElements = {};
            this.sidebarCommentElements = {};
            this.revisionCardElements = new Map();
            this.sidebarContainer = null;
            this.changeAuthorIndex = new Map();
            this.changeElements = [];
            this.changeMeta = [];
            this.moveElements = new Map();
            this.footnoteRefCount = 0;
            this.endnoteRefCount = 0;
            if (this.options.renderComments && this.useHighlight && globalThis.Highlight) {
                this.commentHighlight = new Highlight();
            }
            this.fontAltNames = new Map();
            if (document.fontTablePart?.fonts) {
                for (const f of document.fontTablePart.fonts) {
                    if (f.name && f.altName) {
                        this.fontAltNames.set(f.name, f.altName);
                    }
                }
            }
            const result = [...this.renderDefaultStyle()];
            if (document.themePart) {
                result.push(...this.renderTheme(document.themePart));
            }
            if (document.stylesPart != null) {
                this.styleMap = this.processStyles(document.stylesPart.styles);
                this.applyFontAltNamesToStyles(document.stylesPart.styles);
                result.push(...this.renderStyles(document.stylesPart.styles));
            }
            if (document.numberingPart) {
                this.prodessNumberings(document.numberingPart.domNumberings);
                result.push(...await this.renderNumbering(document.numberingPart.domNumberings));
            }
            if (document.footnotesPart) {
                this.footnoteMap = keyBy(document.footnotesPart.notes, x => x.id);
            }
            if (document.endnotesPart) {
                this.endnoteMap = keyBy(document.endnotesPart.notes, x => x.id);
            }
            if (document.settingsPart) {
                this.defaultTabSize = document.settingsPart.settings?.defaultTabStop;
            }
            if (!options.ignoreFonts && document.fontTablePart)
                result.push(...await this.renderFontTable(document.fontTablePart));
            var sectionElements = this.renderSections(document.documentPart.body);
            if (this.options.inWrapper) {
                if (this.useSidebar) {
                    result.push(this.renderWrapperWithSidebar(sectionElements));
                }
                else {
                    result.push(this.renderWrapper(sectionElements));
                }
            }
            else {
                result.push(...sectionElements);
            }
            if (this.commentHighlight && this.useHighlight) {
                CSS.highlights.set(`${this.className}-comments`, this.commentHighlight);
            }
            else {
                CSS.highlights?.delete(`${this.className}-comments`);
            }
            if (this.showChanges) {
                this.finalizeChangesRendering(result);
            }
            this.postRenderTasks.forEach(t => t());
            await Promise.allSettled(this.tasks);
            this.refreshTabStops();
            return result;
        }
        renderTheme(themePart) {
            const variables = {};
            const fontScheme = themePart.theme?.fontScheme;
            if (fontScheme) {
                if (fontScheme.majorFont?.latinTypeface) {
                    variables['--docx-majorHAnsi-font'] = sanitizeFontFamily(fontScheme.majorFont.latinTypeface);
                }
                if (fontScheme.minorFont?.latinTypeface) {
                    variables['--docx-minorHAnsi-font'] = sanitizeFontFamily(fontScheme.minorFont.latinTypeface);
                }
            }
            const colorScheme = themePart.theme?.colorScheme;
            if (colorScheme) {
                for (const [k, v] of Object.entries(colorScheme.colors)) {
                    if (!isSafeCssIdent(k))
                        continue;
                    const color = sanitizeCssColor(v);
                    if (!color)
                        continue;
                    variables[`--docx-${k}-color`] = color;
                }
            }
            const cssText = this.styleToString(`.${this.className}`, variables);
            return [
                this.h({ tagName: "#comment", children: ["docxjs document theme values"] }),
                this.h({ tagName: "style", children: [cssText] })
            ];
        }
        async renderFontTable(fontsPart) {
            const result = [];
            for (let f of fontsPart.fonts) {
                for (let ref of f.embedFontRefs) {
                    try {
                        const fontData = await this.document.loadFont(ref.id, ref.key);
                        const cssValues = {
                            'font-family': encloseFontFamily(f.name),
                            'src': `url(${fontData})`
                        };
                        if (ref.type == "bold" || ref.type == "boldItalic") {
                            cssValues['font-weight'] = 'bold';
                        }
                        if (ref.type == "italic" || ref.type == "boldItalic") {
                            cssValues['font-style'] = 'italic';
                        }
                        result.push(this.h({ tagName: "#comment", children: [`docxjs ${f.name} font`] }));
                        result.push(this.h({ tagName: "style", children: [this.styleToString(`@font-face`, cssValues)] }));
                    }
                    catch (e) {
                        if (this.options.debug)
                            console.warn(`Can't load font with id ${ref.id} and key ${ref.key}`);
                    }
                }
            }
            return result;
        }
        processStyleName(className) {
            return className ? `${this.className}_${escapeClassName(className)}` : this.className;
        }
        processStyles(styles) {
            const stylesMap = keyBy(styles.filter(x => x.id != null), x => x.id);
            for (const style of styles.filter(x => x.basedOn)) {
                var baseStyle = stylesMap[style.basedOn];
                if (baseStyle) {
                    style.paragraphProps = mergeDeep(style.paragraphProps, baseStyle.paragraphProps);
                    style.runProps = mergeDeep(style.runProps, baseStyle.runProps);
                    for (const baseValues of baseStyle.styles) {
                        const styleValues = style.styles.find(x => x.target == baseValues.target);
                        if (styleValues) {
                            this.copyStyleProperties(baseValues.values, styleValues.values);
                        }
                        else {
                            style.styles.push({ ...baseValues, values: { ...baseValues.values } });
                        }
                    }
                }
                else if (this.options.debug)
                    console.warn(`Can't find base style ${style.basedOn}`);
            }
            for (let style of styles) {
                style.cssName = this.processStyleName(style.id);
            }
            return stylesMap;
        }
        applyFontAltNamesToStyles(styles) {
            if (this.fontAltNames.size === 0)
                return;
            for (const s of styles) {
                for (const sub of (s.styles ?? [])) {
                    if (sub.values)
                        this.applyFontAltNames(sub.values);
                }
            }
        }
        applyFontAltNames(values) {
            if (!values)
                return;
            const existing = values["font-family"];
            if (!existing)
                return;
            const first = existing.split(',')[0].trim().replace(/^['"]|['"]$/g, '');
            const alt = this.fontAltNames.get(first);
            if (!alt)
                return;
            const encodedAlt = sanitizeFontFamily(alt);
            if (existing.includes(encodedAlt))
                return;
            values["font-family"] = `${existing}, ${encodedAlt}`;
        }
        prodessNumberings(numberings) {
            for (let num of numberings.filter(n => n.pStyleName)) {
                const style = this.findStyle(num.pStyleName);
                if (style?.paragraphProps?.numbering) {
                    style.paragraphProps.numbering.level = num.level;
                }
            }
        }
        processElement(element) {
            if (element.children) {
                for (var e of element.children) {
                    e.parent = element;
                    if (e.type == DomType.Table) {
                        this.processTable(e);
                    }
                    else {
                        this.processElement(e);
                    }
                }
            }
        }
        processTable(table) {
            for (var r of table.children) {
                for (var c of r.children) {
                    c.cssStyle = this.copyStyleProperties(table.cellStyle, c.cssStyle, [
                        "border-left", "border-right", "border-top", "border-bottom",
                        "padding-left", "padding-right", "padding-top", "padding-bottom"
                    ]);
                    this.processElement(c);
                }
            }
        }
        copyStyleProperties(input, output, attrs = null) {
            if (!input)
                return output;
            if (output == null)
                output = {};
            if (attrs == null)
                attrs = Object.getOwnPropertyNames(input);
            for (var key of attrs) {
                if (input.hasOwnProperty(key) && !output.hasOwnProperty(key))
                    output[key] = input[key];
            }
            return output;
        }
        createPageElement(className, props, docStyle, pageIndex = 0) {
            const style = { ...docStyle };
            if (props) {
                if (props.pageMargins) {
                    let { left, right } = props.pageMargins;
                    if (props.mirrorMargins && pageIndex % 2 === 1) {
                        [left, right] = [right, left];
                    }
                    style.paddingLeft = left;
                    style.paddingRight = right;
                    style.paddingTop = props.pageMargins.top;
                    style.paddingBottom = props.pageMargins.bottom;
                }
                if (props.pageSize) {
                    if (!this.options.ignoreWidth && !this.options.responsive)
                        style.width = props.pageSize.width;
                    if (!this.options.ignoreHeight)
                        style.minHeight = props.pageSize.height;
                }
                if (props.pageBorders) {
                    for (const edge of ["top", "right", "bottom", "left"]) {
                        const border = props.pageBorders[edge];
                        const css = this.borderToCss(border);
                        if (css) {
                            const key = `border${edge.charAt(0).toUpperCase()}${edge.slice(1)}`;
                            style[key] = css;
                        }
                    }
                }
            }
            const section = this.h({ tagName: "section", className, style });
            addSharedClass(section, "page");
            const orient = props?.pageSize?.orientation;
            if (typeof orient === "string" && /^(landscape|portrait)$/.test(orient)) {
                section.dataset.pageOrientation = orient;
                section.dataset.orientation = orient;
                section.classList.add(`page-${orient}`);
            }
            else if (props?.pageSize) {
                section.dataset.pageOrientation = "portrait";
                section.dataset.orientation = "portrait";
                section.classList.add("page-portrait");
            }
            return section;
        }
        borderToCss(border) {
            if (!border || !border.type || border.type === "none" || border.type === "nil") {
                return null;
            }
            const size = border.size || "0.5pt";
            const styleMap = {
                single: "solid", thick: "solid", double: "double",
                dotted: "dotted", dashed: "dashed", dashSmallGap: "dashed",
                dotDash: "dashed", dotDotDash: "dashed",
                wave: "solid", doubleWave: "double",
            };
            const cssStyle = styleMap[border.type] ?? "solid";
            const color = sanitizeCssColor(border.color) ?? "currentColor";
            return `${size} ${cssStyle} ${color}`;
        }
        createSectionContent(props) {
            const style = {};
            const classNames = [];
            const extraChildren = [];
            if (props.columns && props.columns.numberOfColumns) {
                const { columns } = props;
                const perColumnWidths = columns.columns
                    ?.map(c => c.width)
                    .filter((w) => !!w);
                if (columns.equalWidth === false && perColumnWidths && perColumnWidths.length > 0) {
                    style.display = "grid";
                    style.gridTemplateColumns = perColumnWidths.join(" ");
                    if (columns.space) {
                        style.columnGap = columns.space;
                    }
                }
                else {
                    style.columnCount = `${columns.numberOfColumns}`;
                    style.columnGap = columns.space;
                }
                if (props.columns.separator) {
                    style.columnRule = "1px solid black";
                }
            }
            if (props.docGrid && props.docGrid.linePitch > 0) {
                const pitchPt = props.docGrid.linePitch / 20;
                if (isFinite(pitchPt) && pitchPt > 0) {
                    style.lineHeight = `${pitchPt}pt`;
                }
            }
            if (props.lineNumbering && props.lineNumbering.countBy > 0) {
                const countBy = Math.max(1, Math.floor(Number(props.lineNumbering.countBy)) || 1);
                const restart = props.lineNumbering.restart || "newPage";
                if (restart !== "continuous") {
                    const start = Math.max(0, (props.lineNumbering.start ?? 1) - 1);
                    style.counterReset = `docx-line ${start}`;
                }
                const seq = ++this.lineNumberingArticleSeq;
                const scopeClass = `${this.className}-lnno-${seq}`;
                classNames.push(scopeClass);
                const rule = `.${scopeClass} > p { counter-increment: docx-line; }\n` +
                    `.${scopeClass} > p:nth-of-type(${countBy}n)::before { ` +
                    `content: counter(docx-line); display: inline-block; ` +
                    `width: 2.5em; margin-left: -3em; margin-right: 0.5em; ` +
                    `text-align: right; color: #666; font-size: 0.85em; ` +
                    `vertical-align: top; }`;
                extraChildren.push(this.h({ tagName: "style", children: [rule] }));
            }
            const className = classNames.length ? classNames.join(" ") : undefined;
            return this.h({ tagName: "article", className, style, children: extraChildren });
        }
        renderSections(document) {
            const result = [];
            this.processElement(document);
            const sections = this.splitBySection(document.children, document.props);
            const pages = this.groupByPageBreaks(sections);
            let prevProps = null;
            for (let i = 0, l = pages.length; i < l; i++) {
                this.currentFootnoteIds = [];
                const section = pages[i][0];
                let props = section.sectProps;
                const pageElement = this.createPageElement(this.className, props, document.cssStyle, result.length);
                this.options.renderHeaders && this.renderHeaderFooter(props.headerRefs, props, result.length, prevProps != props, pageElement);
                for (const sect of pages[i]) {
                    var contentElement = this.createSectionContent(sect.sectProps);
                    this.currentSectProps = sect.sectProps;
                    this.renderElements(sect.elements, contentElement);
                    pageElement.appendChild(contentElement);
                    props = sect.sectProps;
                }
                this.currentSectProps = null;
                if (this.options.renderFootnotes) {
                    const notes = this.renderNotes(this.currentFootnoteIds, this.footnoteMap);
                    notes && pageElement.appendChild(notes);
                }
                if (this.options.renderEndnotes && i == l - 1) {
                    const notes = this.renderNotes(this.currentEndnoteIds, this.endnoteMap);
                    notes && pageElement.appendChild(notes);
                }
                this.options.renderFooters && this.renderHeaderFooter(props.footerRefs, props, result.length, prevProps != props, pageElement);
                result.push(pageElement);
                prevProps = props;
            }
            return result;
        }
        renderHeaderFooter(refs, props, page, firstOfSection, into) {
            if (!refs)
                return;
            const evenAndOdd = this.document?.settingsPart?.settings?.evenAndOddHeaders === true;
            var ref = (props.titlePage && firstOfSection ? refs.find(x => x.type == "first") : null)
                ?? (evenAndOdd && page % 2 == 1 ? refs.find(x => x.type == "even") : null)
                ?? refs.find(x => x.type == "default");
            var part = ref && this.document.findPartByRelId(ref.id, this.document.documentPart);
            if (part) {
                this.currentPart = part;
                if (!this.usedHederFooterParts.includes(part.path)) {
                    this.processElement(part.rootElement);
                    this.usedHederFooterParts.push(part.path);
                }
                const [el] = this.renderElements([part.rootElement], into);
                if (props?.pageMargins) {
                    if (part.rootElement.type === DomType.Header) {
                        el.style.marginTop = `calc(${props.pageMargins.header} - ${props.pageMargins.top})`;
                        el.style.minHeight = `calc(${props.pageMargins.top} - ${props.pageMargins.header})`;
                    }
                    else if (part.rootElement.type === DomType.Footer) {
                        el.style.marginBottom = `calc(${props.pageMargins.footer} - ${props.pageMargins.bottom})`;
                        el.style.minHeight = `calc(${props.pageMargins.bottom} - ${props.pageMargins.footer})`;
                    }
                }
                this.currentPart = null;
            }
        }
        isPageBreakElement(elem) {
            if (elem.type != DomType.Break)
                return false;
            if (elem.break == "lastRenderedPageBreak")
                return !this.options.ignoreLastRenderedPageBreak;
            return elem.break == "page";
        }
        isPageBreakSection(prev, next) {
            if (!prev)
                return false;
            if (!next)
                return false;
            return prev.pageSize?.orientation != next.pageSize?.orientation
                || prev.pageSize?.width != next.pageSize?.width
                || prev.pageSize?.height != next.pageSize?.height;
        }
        splitBySection(elements, defaultProps) {
            var current = { sectProps: null, elements: [], pageBreak: false };
            var result = [current];
            for (let elem of elements) {
                if (elem.type == DomType.Paragraph) {
                    const s = this.findStyle(elem.styleName);
                    if (s?.paragraphProps?.pageBreakBefore) {
                        current.sectProps = sectProps;
                        current.pageBreak = true;
                        current = { sectProps: null, elements: [], pageBreak: false };
                        result.push(current);
                    }
                }
                current.elements.push(elem);
                if (elem.type == DomType.Paragraph) {
                    const p = elem;
                    var sectProps = p.sectionProps;
                    var pBreakIndex = -1;
                    var rBreakIndex = -1;
                    if (this.options.breakPages && p.children) {
                        pBreakIndex = p.children.findIndex(r => {
                            rBreakIndex = r.children?.findIndex(this.isPageBreakElement.bind(this)) ?? -1;
                            return rBreakIndex != -1;
                        });
                    }
                    if (sectProps || pBreakIndex != -1) {
                        current.sectProps = sectProps;
                        current.pageBreak = pBreakIndex != -1;
                        current = { sectProps: null, elements: [], pageBreak: false };
                        result.push(current);
                    }
                    if (pBreakIndex != -1) {
                        let breakRun = p.children[pBreakIndex];
                        let splitRun = rBreakIndex < breakRun.children.length - 1;
                        if (pBreakIndex < p.children.length - 1 || splitRun) {
                            var children = elem.children;
                            var newParagraph = { ...elem, children: children.slice(pBreakIndex) };
                            elem.children = children.slice(0, pBreakIndex);
                            current.elements.push(newParagraph);
                            if (splitRun) {
                                let runChildren = breakRun.children;
                                let newRun = { ...breakRun, children: runChildren.slice(0, rBreakIndex) };
                                elem.children.push(newRun);
                                breakRun.children = runChildren.slice(rBreakIndex);
                            }
                        }
                    }
                }
            }
            let currentSectProps = null;
            for (let i = result.length - 1; i >= 0; i--) {
                if (result[i].sectProps == null) {
                    result[i].sectProps = currentSectProps ?? defaultProps;
                }
                else {
                    currentSectProps = result[i].sectProps;
                }
            }
            return result;
        }
        groupByPageBreaks(sections) {
            let current = [];
            let prev;
            const result = [current];
            for (let s of sections) {
                current.push(s);
                if (this.options.ignoreLastRenderedPageBreak || s.pageBreak || this.isPageBreakSection(prev, s.sectProps))
                    result.push(current = []);
                prev = s.sectProps;
            }
            return result.filter(x => x.length > 0);
        }
        renderWrapper(children) {
            const wrapperChildren = [...this.renderProtectionBadge(), ...children];
            const wrapper = this.h({ tagName: "div", className: `${this.className}-wrapper`, children: wrapperChildren });
            addSharedClass(wrapper, "wrapper");
            wrapper.setAttribute("role", "document");
            this.applyDocumentPropsDataset(wrapper);
            return wrapper;
        }
        renderProtectionBadge() {
            if (!this.options.showProtectionBadge)
                return [];
            const protection = this.document?.settingsPart?.settings?.documentProtection;
            if (!protection)
                return [];
            const labels = {
                readOnly: 'Protected document (read-only)',
                trackedChanges: 'Protected document (tracked changes required)',
                comments: 'Protected document (comments only)',
                forms: 'Protected document (form fields only)',
                none: 'Protected document',
            };
            const label = labels[protection.edit ?? 'none'] ?? 'Protected document';
            const badge = this.h({ tagName: "div", className: `${this.className}-protected-badge`, children: [label] });
            badge.setAttribute("role", "status");
            if (protection.edit)
                badge.dataset.protectionEdit = protection.edit;
            if (protection.enforcement)
                badge.dataset.protectionEnforced = "1";
            return [badge];
        }
        applyDocumentPropsDataset(wrapper) {
            if (!this.options.emitDocumentProps)
                return;
            const core = this.document?.corePropsPart?.props;
            if (!core)
                return;
            if (core.title)
                wrapper.dataset.docTitle = core.title;
            if (core.subject)
                wrapper.dataset.docSubject = core.subject;
            if (core.creator)
                wrapper.dataset.docAuthor = core.creator;
            if (core.lastModifiedBy)
                wrapper.dataset.docLastModifiedBy = core.lastModifiedBy;
            if (core.description)
                wrapper.dataset.docDescription = core.description;
            if (core.keywords)
                wrapper.dataset.docKeywords = core.keywords;
            if (core.created)
                wrapper.dataset.docCreated = core.created;
            if (core.modified)
                wrapper.dataset.docModified = core.modified;
        }
        renderWrapperWithSidebar(sectionElements) {
            const c = this.className;
            const docContainer = this.h({ tagName: "div", className: `${c}-doc-container`, children: sectionElements });
            this.sidebarContainer = this.h({
                tagName: "div",
                className: `${c}-comment-sidebar ${c}-sidebar-${this.sidebarLayout}`
            });
            const contentArea = this.h({
                tagName: "div",
                className: `${c}-sidebar-content`,
                children: []
            });
            this.sidebarContainer.appendChild(contentArea);
            this.renderSidebarComments(contentArea);
            const wrapper = this.h({
                tagName: "div",
                className: `${c}-wrapper`,
                children: [...this.renderProtectionBadge(), docContainer, this.sidebarContainer]
            });
            wrapper.setAttribute("role", "document");
            addSharedClass(wrapper, "wrapper");
            this.applyDocumentPropsDataset(wrapper);
            this.later(() => {
                this.setupSidebarScrollSync(docContainer, contentArea, wrapper);
            });
            return wrapper;
        }
        setupSidebarScrollSync(docContainer, sidebarContent, wrapper) {
            if (this.sidebarLayout === 'packed')
                return;
            const CARD_GAP = 8;
            const positionCards = () => {
                const anchored = [];
                for (const [id, sidebarEl] of Object.entries(this.sidebarCommentElements)) {
                    if (!sidebarEl.isConnected)
                        continue;
                    const anchor = this.commentAnchorElements[id]?.[0];
                    if (!anchor?.isConnected)
                        continue;
                    anchored.push({ el: sidebarEl, anchor, desiredTop: 0 });
                }
                for (const meta of this.changeMeta) {
                    const card = this.revisionCardElements.get(meta.id ?? '');
                    if (!card?.isConnected || !meta.el.isConnected)
                        continue;
                    anchored.push({ el: card, anchor: meta.el, desiredTop: 0 });
                }
                if (anchored.length === 0)
                    return;
                const previousPosition = sidebarContent.style.position;
                if (previousPosition !== 'relative' && previousPosition !== 'absolute') {
                    sidebarContent.style.position = 'relative';
                }
                for (const { el } of anchored) {
                    el.style.marginTop = '';
                    el.style.position = 'absolute';
                    el.style.top = '0';
                    el.style.left = '0';
                    el.style.right = '0';
                }
                const sidebarRect = sidebarContent.getBoundingClientRect();
                for (const entry of anchored) {
                    const r = entry.anchor.getBoundingClientRect();
                    entry.desiredTop = r.top - sidebarRect.top + sidebarContent.scrollTop;
                }
                anchored.sort((a, b) => a.desiredTop - b.desiredTop);
                let floor = -Infinity;
                let maxBottom = 0;
                for (const entry of anchored) {
                    const target = Math.max(entry.desiredTop, floor);
                    entry.el.style.top = `${target}px`;
                    const bottom = target + entry.el.offsetHeight;
                    floor = bottom + CARD_GAP;
                    if (bottom > maxBottom)
                        maxBottom = bottom;
                }
                sidebarContent.style.minHeight = `${maxBottom}px`;
            };
            let rafId;
            const schedule = () => {
                cancelAnimationFrame(rafId);
                rafId = requestAnimationFrame(positionCards);
            };
            if (typeof ResizeObserver !== "undefined") {
                const ro = new ResizeObserver(schedule);
                if (wrapper)
                    ro.observe(wrapper);
                ro.observe(docContainer);
                for (const el of Object.values(this.sidebarCommentElements)) {
                    if (el.isConnected)
                        ro.observe(el);
                }
            }
            setTimeout(positionCards, 100);
            setTimeout(positionCards, 500);
            setTimeout(positionCards, 1500);
        }
        renderSidebarComments(container) {
            const commentsPart = this.document.commentsPart;
            if (!commentsPart)
                return;
            const comments = commentsPart.topLevelComments.length > 0
                ? commentsPart.topLevelComments
                : commentsPart.comments;
            for (const comment of comments) {
                const el = this.renderSidebarComment(comment, false);
                if (el)
                    container.appendChild(el);
            }
        }
        renderSidebarComment(comment, isReply) {
            const c = this.className;
            const headerChildren = [
                this.h({ tagName: "span", className: `${c}-comment-author`, children: [comment.author ?? "Unknown"] }),
                this.h({ tagName: "span", className: `${c}-comment-date`, children: [comment.date ? new Date(comment.date).toLocaleString() : ""] })
            ];
            if (comment.done) {
                headerChildren.push(this.h({ tagName: "span", className: `${c}-comment-done`, children: ["Done"] }));
            }
            const header = this.h({
                tagName: "div",
                className: `${c}-comment-header`,
                children: headerChildren
            });
            const bodyEl = this.h({
                tagName: "div",
                className: `${c}-comment-body`,
                children: this.renderElements(comment.children)
            });
            const children = [header, bodyEl];
            if (comment.replies && comment.replies.length > 0) {
                const repliesContainer = this.h({
                    tagName: "div",
                    className: `${c}-comment-replies`,
                    children: comment.replies.map(r => this.renderSidebarComment(r, true))
                });
                const threadToggle = this.h({
                    tagName: "button",
                    className: `${c}-thread-toggle`,
                    children: [`${comment.replies.length} ${comment.replies.length === 1 ? 'reply' : 'replies'}`]
                });
                children.push(threadToggle);
                children.push(repliesContainer);
                this.later(() => {
                    threadToggle.addEventListener("click", (ev) => {
                        ev.stopPropagation();
                        repliesContainer.classList.toggle(`${c}-replies-collapsed`);
                        threadToggle.classList.toggle(`${c}-thread-collapsed`);
                    });
                });
            }
            const commentEl = this.h({
                tagName: "div",
                className: cx(`${c}-sidebar-comment`, isReply && `${c}-sidebar-reply`),
                children
            });
            commentEl.dataset.commentId = comment.id;
            if (!isReply) {
                this.sidebarCommentElements[comment.id] = commentEl;
                this.later(() => {
                    commentEl.addEventListener("click", () => {
                        const anchors = this.commentAnchorElements[comment.id];
                        if (anchors && anchors.length > 0) {
                            anchors[0].scrollIntoView({ behavior: "smooth", block: "center" });
                        }
                    });
                });
            }
            return commentEl;
        }
        renderDefaultStyle() {
            var c = this.className;
            var wrapperStyle = `
.${c}-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; } 
.${c}-wrapper>section.${c} { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }`;
            if (this.options.hideWrapperOnPrint) {
                wrapperStyle = `@media not print { ${wrapperStyle} }`;
            }
            var styleText = `${wrapperStyle}
.${c} { color: black; hyphens: auto; text-underline-position: from-font; }
section.${c} { box-sizing: border-box; display: flex; flex-flow: column nowrap; position: relative; overflow: hidden; }
section.${c}>article { margin-bottom: auto; z-index: 1; }
section.${c}>footer { z-index: 1; }
.${c} table { border-collapse: collapse; }
.${c} table td, .${c} table th { vertical-align: top; }
.${c} p { margin: 0pt; min-height: 1em; }
.${c} span { white-space: pre-wrap; overflow-wrap: break-word; }
.${c} a { color: inherit; text-decoration: inherit; }
.${c} svg { fill: transparent; }
.${c}-footnote-ref, .${c}-endnote-ref { font-size: 0.65em; line-height: 0; vertical-align: super; }
/* Footnote / endnote list: hide the default browser list marker and render
 * our own superscript counter before each <li>, so the item number matches
 * the style of the inline footnote reference. The list-item counter
 * automatically respects the <ol> start attribute set by page-break.ts. */
section.${c}>ol { list-style: none; padding-left: 0; }
section.${c}>ol>li { position: relative; padding-left: 1.25em; }
section.${c}>ol>li::before {
    content: counter(list-item);
    position: absolute;
    left: 0;
    font-size: 0.65em;
    line-height: 0;
    vertical-align: super;
    top: 0.35em;
}
/* OLE placeholder — legacy w:object embeds (Equation.3, Excel.Sheet.12,
 * Package, ...) surface as a labelled inline span. Hard-coded label set
 * lives in renderOleObject; the attribute selector here is safe because
 * data-progid is only populated from that allowlist. */
.${c}-ole-placeholder {
    display: inline-block;
    padding: 1px 6px;
    margin: 0 1px;
    border: 1px dashed #9aa0a6;
    background: #f4f5f7;
    color: #5f6368;
    font-size: 0.85em;
    font-style: italic;
    border-radius: 2px;
    white-space: nowrap;
    vertical-align: baseline;
}
/* Protection badge — rendered only when the caller opts in via
 * Options.showProtectionBadge and the DOCX carries w:documentProtection.
 * The label is hard-coded (see renderProtectionBadge) so the CSS here is
 * purely cosmetic and safe to emit unconditionally. */
.${c}-protected-badge {
    align-self: stretch;
    margin: 0 0 12px 0;
    padding: 6px 12px;
    background: #fff6d6;
    color: #5a4400;
    border: 1px solid #e4c66f;
    border-radius: 3px;
    font-family: system-ui, -apple-system, sans-serif;
    font-size: 12px;
    text-align: center;
}
`;
            if (this.options.renderComments) {
                if (this.useSidebar) {
                    styleText += `
.${c}-wrapper { flex-flow: row !important; align-items: flex-start !important; }
.${c}-doc-container { flex: 1; display: flex; flex-flow: column; align-items: center; min-width: 0; overflow: auto; padding: 30px; padding-bottom: 0; }
.${c}-doc-container>section.${c} { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }
.${c}-comment-sidebar { width: 320px; min-width: 260px; display: flex; flex-direction: column; transition: width 0.2s, min-width 0.2s, padding 0.2s; }
/* packed mode: panel stays pinned as a short compact list at the top of the viewport. Background + border frame the compact list. */
.${c}-comment-sidebar.${c}-sidebar-packed { position: sticky; top: 0; height: 100vh; overflow: hidden; align-self: flex-start; background: #fafafa; border-left: 1px solid #ddd; }
/* anchored mode: panel grows to match the document height and rides the same scroll container so each card stays next to its anchor. No background/border — cards float on the page backdrop. */
.${c}-comment-sidebar.${c}-sidebar-anchored { align-self: stretch; background: transparent; border-left: none; }
.${c}-sidebar-packed .${c}-sidebar-content { flex: 1; overflow-y: auto; padding: 8px; }
.${c}-sidebar-anchored .${c}-sidebar-content { padding: 8px; }
.${c}-sidebar-comment { background: white; border: 1px solid #e0e0e0; border-radius: 6px; padding: 10px; margin-bottom: 8px; cursor: pointer; transition: box-shadow 0.2s, border-color 0.2s; }
.${c}-sidebar-comment:hover { border-color: #4a90d9; box-shadow: 0 1px 4px rgba(74, 144, 217, 0.2); }
.${c}-sidebar-reply { margin-left: 16px; border-left: 3px solid #4a90d9; background: #f8fbff; }
.${c}-comment-header { display: flex; align-items: baseline; gap: 8px; margin-bottom: 4px; flex-wrap: wrap; }
.${c}-comment-author { font-weight: 600; font-size: 0.85rem; color: #333; }
.${c}-comment-date { font-size: 0.75rem; color: #999; }
.${c}-comment-done { font-size: 0.7rem; background: #4caf50; color: white; padding: 1px 6px; border-radius: 3px; }
.${c}-comment-body { font-size: 0.85rem; color: #444; margin-bottom: 6px; line-height: 1.4; }
.${c}-comment-body p { margin: 2px 0; }
.${c}-comment-replies { margin-top: 6px; }
.${c}-replies-collapsed { display: none; }
.${c}-thread-toggle { background: none; border: none; color: #4a90d9; cursor: pointer; font-size: 0.8rem; padding: 2px 0; margin-top: 4px; }
.${c}-thread-toggle:hover { text-decoration: underline; }
.${c}-thread-collapsed::before { content: "▶ "; }
.${c}-thread-toggle:not(.${c}-thread-collapsed)::before { content: "▼ "; }
.${c}-comment-focused { border-color: #ff9800 !important; box-shadow: 0 0 8px rgba(255, 152, 0, 0.4) !important; }
.${c}-comment-anchor-start { cursor: pointer; }
::highlight(${c}-comments) { background-color: rgba(255, 212, 0, 0.35); }
.${c}-no-highlight .${c}-comment-anchor-start { cursor: default; }
`;
                }
                else {
                    styleText += `
.${c}-comment-ref { cursor: default; }
.${c}-comment-popover { display: none; z-index: 1000; padding: 0.5rem; background: white; position: absolute; box-shadow: 0 0 0.25rem rgba(0, 0, 0, 0.25); width: 30ch; }
.${c}-comment-ref:hover~.${c}-comment-popover { display: block; }
.${c}-comment-author,.${c}-comment-date { font-size: 0.875rem; color: #888; }
`;
                }
            }
            if (this.showChanges) {
                styleText += this.changesStyles();
            }
            if (this.options.responsive) {
                styleText += `
.${c}-wrapper { padding: 8px; }
section.${c} { width: auto !important; max-width: 100%; min-width: 0; box-sizing: border-box; }
.${c} img { max-width: 100%; height: auto; }
.${c} table { max-width: 100%; table-layout: auto; }
@media (max-width: 768px) {
  .${c}-wrapper { padding: 4px; background: #fff; }
  .${c}-wrapper>section.${c} { box-shadow: none; margin-bottom: 12px; }
  section.${c} { padding-left: 12px !important; padding-right: 12px !important; }
  .${c} [data-drawing-anchor="true"] { float: none !important; display: block !important; position: static !important; width: auto !important; max-width: 100%; margin: 0.5em 0 !important; }
}
`;
            }
            return [
                this.h({ tagName: "#comment", children: ["docxjs library predefined styles"] }),
                this.h({ tagName: "style", children: [styleText] })
            ];
        }
        changesStyles() {
            const c = this.className;
            const palette = [
                "#2563eb", "#dc2626", "#16a34a", "#9333ea",
                "#ea580c", "#0891b2", "#c026d3", "#65a30d"
            ];
            let css = `
.${c} ins { text-decoration: underline; text-decoration-thickness: 2px; background: transparent; }
.${c} del { text-decoration: line-through; text-decoration-thickness: 2px; }
.${c} .${c}-move-from { text-decoration: line-through double; text-decoration-thickness: 1px; cursor: pointer; }
.${c} .${c}-move-to { text-decoration: underline double; text-decoration-thickness: 1px; cursor: pointer; }
.${c} .${c}-formatting-revision { text-decoration: underline dotted; text-decoration-thickness: 1px; cursor: help; }
.${c}-paragraph-mark { margin-left: 2px; font-weight: bold; user-select: none; }
.${c}-paragraph-mark-deleted { text-decoration: line-through; }
.${c}-row-inserted > td { background: color-mix(in srgb, currentColor 8%, transparent); }
.${c}-row-deleted > td { background: color-mix(in srgb, currentColor 10%, transparent); text-decoration: line-through; text-decoration-color: currentColor; text-decoration-thickness: 2px; }
.${c}-revision-kind { margin-left: auto; font-size: 0.7rem; padding: 1px 6px; border: 1px solid currentColor; border-radius: 3px; text-transform: uppercase; }
.${c}-revision-card { border-left: 3px solid currentColor; }
.${c}-change-bar { position: relative; }
.${c}-change-bar::before { content: ""; position: absolute; left: -12px; top: 0; bottom: 0; width: 2px; background: currentColor; opacity: 0.55; }
.${c}-legend { display: flex; flex-wrap: wrap; gap: 12px; align-items: center; padding: 8px 12px; margin: 0 auto 12px; background: #f5f5f5; border: 1px solid #ddd; border-radius: 4px; font-size: 0.85rem; color: #333; max-width: calc(100% - 60px); }
.${c}-legend-label { font-weight: 600; margin-right: 4px; }
.${c}-legend-item { display: inline-flex; align-items: center; gap: 4px; }
.${c}-legend-swatch { display: inline-block; width: 12px; height: 12px; border-radius: 2px; }
`;
            for (let i = 0; i < HtmlRenderer.CHANGE_PALETTE_SIZE; i++) {
                css += `.${c}-change-author-${i} { color: ${palette[i]}; text-decoration-color: ${palette[i]}; }\n`;
            }
            return css;
        }
        async renderNumbering(numberings) {
            var styleText = "";
            var resetCounters = [];
            for (var num of numberings) {
                if (!isSafeCssIdent(String(num.id)) || !Number.isInteger(num.level)) {
                    continue;
                }
                var selector = `p.${this.numberingClass(num.id, num.level)}`;
                var listStyleType = "none";
                if (num.bullet) {
                    if (!isSafeCssIdent(String(num.bullet.src))) {
                        continue;
                    }
                    let valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();
                    styleText += this.styleToString(`${selector}:before`, {
                        "content": "' '",
                        "display": "inline-block",
                        "background": `var(${valiable})`
                    });
                    try {
                        const imgData = await this.document.loadNumberingImage(num.bullet.src);
                        styleText += `${this.rootSelector} { ${valiable}: url(${imgData}) }`;
                    }
                    catch (e) {
                        if (this.options.debug)
                            console.warn(`Can't load numbering image with src ${num.bullet.src}`);
                    }
                }
                else if (num.levelText) {
                    let counter = this.numberingCounter(num.id, num.level);
                    const counterReset = counter + " " + (num.start - 1);
                    const restart = num.restart;
                    const restartDefault = restart === undefined || restart === -1;
                    if (restartDefault) {
                        if (num.level > 0) {
                            styleText += this.styleToString(`p.${this.numberingClass(num.id, num.level - 1)}`, {
                                "counter-set": counterReset
                            });
                        }
                    }
                    else if (Number.isInteger(restart) && restart > 0 && restart <= num.level) {
                        styleText += this.styleToString(`p.${this.numberingClass(num.id, restart - 1)}`, {
                            "counter-set": counterReset
                        });
                    }
                    resetCounters.push(counterReset);
                    const levelFormat = this.numFormatToCssValue(num.format);
                    const beforeStyle = {
                        "content": this.levelTextToContent(num.levelText, num.suff, num.id, levelFormat, num.isLgl === true),
                        "counter-increment": counter,
                        ...num.rStyle,
                    };
                    const justifyMap = {
                        left: "left",
                        right: "right",
                        center: "center",
                        start: "start",
                        end: "end",
                    };
                    const justify = num.justification && justifyMap[num.justification];
                    if (justify) {
                        beforeStyle["text-align"] = justify;
                    }
                    styleText += this.styleToString(`${selector}:before`, beforeStyle);
                }
                else {
                    listStyleType = this.numFormatToCssValue(num.format);
                }
                styleText += this.styleToString(selector, {
                    "display": "list-item",
                    "list-style-position": "inside",
                    "list-style-type": listStyleType,
                    ...num.pStyle
                });
            }
            if (resetCounters.length > 0) {
                styleText += this.styleToString(this.rootSelector, {
                    "counter-reset": resetCounters.join(" ")
                });
            }
            return [
                this.h({ tagName: "#comment", children: ["docxjs document numbering styles"] }),
                this.h({ tagName: "style", children: [styleText] })
            ];
        }
        renderStyles(styles) {
            var styleText = "";
            const stylesMap = this.styleMap;
            const defautStyles = keyBy(styles.filter(s => s.isDefault), s => s.target);
            for (const style of styles) {
                var subStyles = style.styles;
                if (style.linked) {
                    var linkedStyle = style.linked && stylesMap[style.linked];
                    if (linkedStyle)
                        subStyles = subStyles.concat(linkedStyle.styles);
                    else if (this.options.debug)
                        console.warn(`Can't find linked style ${style.linked}`);
                }
                for (const subStyle of subStyles) {
                    var selector = `${style.target ?? ''}.${style.cssName}`;
                    if (style.target != subStyle.target)
                        selector += ` ${subStyle.target}`;
                    if (defautStyles[style.target] == style)
                        selector = `.${this.className} ${style.target}, ` + selector;
                    this.applyThemeColorSideband(subStyle.values);
                    styleText += this.styleToString(selector, subStyle.values);
                }
            }
            return [
                this.h({ tagName: "#comment", children: ["docxjs document styles"] }),
                this.h({ tagName: "style", children: [styleText] })
            ];
        }
        renderNotes(noteIds, notesMap) {
            const seenIds = new Set();
            const uniqueIds = [];
            for (const id of noteIds) {
                if (!seenIds.has(id)) {
                    seenIds.add(id);
                    uniqueIds.push(id);
                }
            }
            var notes = uniqueIds.map(id => notesMap[id]).filter(x => x);
            if (notes.length > 0) {
                const renderedChildren = this.renderElements(notes);
                for (let i = 0; i < notes.length && i < renderedChildren.length; i++) {
                    const node = renderedChildren[i];
                    const id = notes[i]?.id;
                    if (node && typeof node.setAttribute === 'function' && id) {
                        node.setAttribute('data-footnote-id', id);
                        node.setAttribute('data-footnote', '');
                    }
                }
                return this.h({ tagName: "ol", children: renderedChildren });
            }
        }
        renderElement(elem) {
            switch (elem.type) {
                case DomType.Paragraph:
                    return this.renderParagraph(elem);
                case DomType.BookmarkStart:
                    return this.renderBookmarkStart(elem);
                case DomType.BookmarkEnd:
                    return null;
                case DomType.Run:
                    return this.renderRun(elem);
                case DomType.Table:
                    return this.renderTable(elem);
                case DomType.Row:
                    return this.renderTableRow(elem);
                case DomType.Cell:
                    return this.renderTableCell(elem);
                case DomType.Hyperlink:
                    return this.renderHyperlink(elem);
                case DomType.SmartTag:
                    return this.renderSmartTag(elem);
                case DomType.SimpleField:
                    return this.renderSimpleField(elem);
                case DomType.ComplexField:
                case DomType.Instruction:
                    return null;
                case DomType.Drawing:
                    return this.renderDrawing(elem);
                case DomType.Image:
                    return this.renderImage(elem);
                case DomType.DrawingShape:
                    return this.renderDrawingShape(elem);
                case DomType.DrawingGroup:
                    return this.renderDrawingShapeGroup(elem);
                case DomType.Chart:
                    return this.renderChart(elem);
                case DomType.ChartEx:
                    return this.renderChartEx(elem);
                case DomType.SmartArt:
                    return this.renderSmartArtPlaceholder(elem);
                case DomType.Text:
                    return this.renderText(elem);
                case DomType.Text:
                    return this.renderText(elem);
                case DomType.DeletedText:
                    return this.renderDeletedText(elem);
                case DomType.Tab:
                    return this.renderTab(elem);
                case DomType.Symbol:
                    return this.renderSymbol(elem);
                case DomType.Break:
                    return this.renderBreak(elem);
                case DomType.Footer:
                    return this.renderContainer(elem, "footer");
                case DomType.Header:
                    return this.renderContainer(elem, "header");
                case DomType.Footnote:
                case DomType.Endnote:
                    return this.renderContainer(elem, "li");
                case DomType.FootnoteReference:
                    return this.renderFootnoteReference(elem);
                case DomType.EndnoteReference:
                    return this.renderEndnoteReference(elem);
                case DomType.NoBreakHyphen:
                    return this.h({ tagName: "wbr" });
                case DomType.VmlPicture:
                    return this.renderVmlPicture(elem);
                case DomType.VmlElement:
                    return this.renderVmlElement(elem);
                case DomType.OleObject:
                    return this.renderOleObject(elem);
                case DomType.MmlMath:
                    return this.renderContainerNS(elem, ns.mathML, "math", { xmlns: ns.mathML });
                case DomType.MmlMathParagraph:
                    return this.renderContainer(elem, "span");
                case DomType.MmlFraction:
                    return this.renderContainerNS(elem, ns.mathML, "mfrac");
                case DomType.MmlBase:
                    return this.renderContainerNS(elem, ns.mathML, elem.parent.type == DomType.MmlMatrixRow ? "mtd" : "mrow");
                case DomType.MmlNumerator:
                case DomType.MmlDenominator:
                case DomType.MmlFunction:
                case DomType.MmlLimit:
                case DomType.MmlBox:
                    return this.renderContainerNS(elem, ns.mathML, "mrow");
                case DomType.MmlGroupChar:
                    return this.renderMmlGroupChar(elem);
                case DomType.MmlLimitLower:
                    return this.renderContainerNS(elem, ns.mathML, "munder");
                case DomType.MmlMatrix:
                    return this.renderContainerNS(elem, ns.mathML, "mtable");
                case DomType.MmlMatrixRow:
                    return this.renderContainerNS(elem, ns.mathML, "mtr");
                case DomType.MmlRadical:
                    return this.renderMmlRadical(elem);
                case DomType.MmlSuperscript:
                    return this.renderContainerNS(elem, ns.mathML, "msup");
                case DomType.MmlSubscript:
                    return this.renderContainerNS(elem, ns.mathML, "msub");
                case DomType.MmlDegree:
                case DomType.MmlSuperArgument:
                case DomType.MmlSubArgument:
                    return this.renderContainerNS(elem, ns.mathML, "mn");
                case DomType.MmlFunctionName:
                    return this.renderContainerNS(elem, ns.mathML, "ms");
                case DomType.MmlDelimiter:
                    return this.renderMmlDelimiter(elem);
                case DomType.MmlRun:
                    return this.renderMmlRun(elem);
                case DomType.MmlNary:
                    return this.renderMmlNary(elem);
                case DomType.MmlPreSubSuper:
                    return this.renderMmlPreSubSuper(elem);
                case DomType.MmlBar:
                    return this.renderMmlBar(elem);
                case DomType.MmlEquationArray:
                    return this.renderMllList(elem);
                case DomType.MmlAccent:
                    return this.renderMmlAccent(elem);
                case DomType.MmlBorderBox:
                    return this.renderMmlBorderBox(elem);
                case DomType.MmlSubSuperscript:
                    return this.renderMmlSubSuperscript(elem);
                case DomType.MmlPhantom:
                    return this.renderContainerNS(elem, ns.mathML, "mphantom");
                case DomType.MmlGroup:
                    return this.renderMmlGroup(elem);
                case DomType.Inserted:
                    return this.renderInserted(elem);
                case DomType.Deleted:
                    return this.renderDeleted(elem);
                case DomType.MoveFrom:
                    return this.renderMoveFrom(elem);
                case DomType.MoveTo:
                    return this.renderMoveTo(elem);
                case DomType.CommentRangeStart:
                    return this.renderCommentRangeStart(elem);
                case DomType.CommentRangeEnd:
                    return this.renderCommentRangeEnd(elem);
                case DomType.CommentReference:
                    return this.renderCommentReference(elem);
                case DomType.AltChunk:
                    return null;
                case DomType.Sdt:
                    return this.renderSdt(elem);
                case DomType.Ruby:
                    return this.renderRuby(elem);
                case DomType.RubyBase:
                    return this.renderContainer(elem, "span");
                case DomType.RubyText:
                    return this.renderContainer(elem, "rt");
                case DomType.FitText:
                    return this.renderFitText(elem);
                case DomType.BidiOverride:
                    return this.renderBidiOverride(elem);
            }
            return null;
        }
        renderElements(elems, into) {
            if (elems == null)
                return null;
            const hasComplexField = elems.some(e => isComplexFieldBeginRun(e));
            const source = hasComplexField ? this.groupComplexFields(elems) : elems;
            var result = source.flatMap(e => {
                if (e instanceof Node)
                    return [e];
                return this.renderElement(e);
            }).filter(e => e != null);
            if (into)
                result.forEach(c => into.appendChild(isString(c) ? document.createTextNode(c) : c));
            return result;
        }
        groupComplexFields(elems) {
            const stack = [];
            const topLevel = [];
            for (let i = 0; i < elems.length; i++) {
                const el = elems[i];
                const ct = complexFieldCharType(el);
                if (ct === 'begin') {
                    stack.push({ beginIdx: i, instrText: [], sepIdx: -1, endIdx: -1, nested: [] });
                    continue;
                }
                if (ct === 'separate') {
                    if (stack.length > 0 && stack[stack.length - 1].sepIdx === -1) {
                        stack[stack.length - 1].sepIdx = i;
                    }
                    continue;
                }
                if (ct === 'end') {
                    if (stack.length === 0)
                        continue;
                    const closed = stack.pop();
                    closed.endIdx = i;
                    if (stack.length > 0) {
                        stack[stack.length - 1].nested.push(closed);
                    }
                    else {
                        topLevel.push(closed);
                    }
                    continue;
                }
                if (stack.length > 0 && stack[stack.length - 1].sepIdx === -1) {
                    if (el && el.type === DomType.Run && el.children) {
                        for (const c of el.children) {
                            if (c.type === DomType.Instruction) {
                                stack[stack.length - 1].instrText.push(c.text ?? '');
                            }
                        }
                    }
                }
            }
            for (const unterminated of stack) {
                topLevel.push(unterminated);
            }
            topLevel.sort((a, b) => a.beginIdx - b.beginIdx);
            const groupByBegin = new Map();
            for (const g of topLevel)
                groupByBegin.set(g.beginIdx, g);
            const out = [];
            let i = 0;
            while (i < elems.length) {
                const g = groupByBegin.get(i);
                if (g) {
                    const rendered = this.renderComplexFieldGroup(g, elems);
                    out.push(...rendered);
                    i = g.endIdx === -1 ? elems.length : g.endIdx + 1;
                    continue;
                }
                out.push(elems[i]);
                i++;
            }
            return out;
        }
        renderComplexFieldGroup(group, elems) {
            const instruction = group.instrText.join('');
            const parsed = parseFieldInstruction(instruction);
            const unterminated = group.endIdx === -1;
            if (unterminated && group.sepIdx === -1) {
                return [];
            }
            const resultStart = group.sepIdx + 1;
            const resultEnd = unterminated ? elems.length : group.endIdx;
            if (resultEnd <= resultStart) {
                return unterminated ? [this.h({ tagName: '#comment', children: ['docxjs: unterminated field'] })] : [];
            }
            const nestedByBegin = new Map();
            for (const n of group.nested)
                nestedByBegin.set(n.beginIdx, n);
            const rendered = [];
            let j = resultStart;
            while (j < resultEnd) {
                const inner = nestedByBegin.get(j);
                if (inner) {
                    const innerNodes = this.renderComplexFieldGroup(inner, elems);
                    rendered.push(...innerNodes);
                    j = inner.endIdx === -1 ? resultEnd : inner.endIdx + 1;
                    continue;
                }
                const r = elems[j];
                let n = null;
                if (r && r.type === DomType.Run) {
                    n = this.renderRun(r, true);
                }
                else if (r) {
                    n = this.renderElement(r);
                }
                if (n != null) {
                    if (Array.isArray(n))
                        rendered.push(...n);
                    else
                        rendered.push(n);
                }
                j++;
            }
            const beginRun = elems[group.beginIdx];
            let ffData;
            if (beginRun && beginRun.type === DomType.Run && beginRun.children) {
                for (const c of beginRun.children) {
                    if (c.type === DomType.ComplexField) {
                        const fc = c;
                        if (fc.charType === 'begin' && fc.ffData) {
                            ffData = fc.ffData;
                            break;
                        }
                    }
                }
            }
            const wrapped = this.wrapFieldResult(rendered, parsed, ffData);
            if (unterminated) {
                return [this.h({ tagName: '#comment', children: ['docxjs: unterminated field'] }), ...wrapped];
            }
            return wrapped;
        }
        renderSimpleField(elem) {
            const parsed = parseFieldInstruction(elem.instruction);
            const children = this.renderElements(elem.children) ?? [];
            return this.wrapFieldResult(children, parsed);
        }
        wrapFieldResult(children, parsed, ffData) {
            const code = parsed.code;
            if (code === 'HYPERLINK') {
                if (!children || children.length === 0)
                    return children ?? [];
                return this.wrapHyperlinkFieldResult(children, parsed);
            }
            if (code === 'REF' || code === 'PAGEREF') {
                if (!children || children.length === 0)
                    return children ?? [];
                const anchor = parsed.args[0] ?? '';
                if (!anchor)
                    return children;
                const a = this.h({ tagName: "a" });
                a.setAttribute('href', '#' + anchor);
                a.dataset.field = code;
                a.classList.add(`${this.className}-field-${code.toLowerCase()}`);
                a.classList.add('field-ref');
                children.forEach(c => a.appendChild(c));
                return [a];
            }
            if (code === 'FORMTEXT') {
                return this.renderLegacyFormTextField(children, ffData);
            }
            if (code === 'FORMCHECKBOX') {
                return this.renderLegacyFormCheckboxField(ffData);
            }
            if (code === 'FORMDROPDOWN') {
                return this.renderLegacyFormDropdownField(children, ffData);
            }
            if (!children || children.length === 0)
                return children ?? [];
            return children;
        }
        renderLegacyFormTextField(children, ffData) {
            const value = children
                .map(n => (n instanceof Node ? (n.textContent ?? '') : String(n)))
                .join('');
            const input = this.h({ tagName: "input" });
            input.setAttribute("type", "text");
            input.setAttribute("disabled", "");
            input.setAttribute("value", value);
            if (ffData?.maxLength != null && Number.isFinite(ffData.maxLength) && ffData.maxLength > 0) {
                input.setAttribute("maxlength", String(Math.min(10000, ffData.maxLength)));
            }
            input.dataset.field = 'FORMTEXT';
            input.classList.add(`${this.className}-field-formtext`);
            return [input];
        }
        renderLegacyFormCheckboxField(ffData) {
            const box = this.h({ tagName: "input" });
            box.setAttribute("type", "checkbox");
            box.setAttribute("disabled", "");
            if (ffData?.checked)
                box.setAttribute("checked", "");
            box.dataset.field = 'FORMCHECKBOX';
            box.classList.add(`${this.className}-field-formcheckbox`);
            return [box];
        }
        renderLegacyFormDropdownField(children, ffData) {
            const select = this.h({ tagName: "select" });
            select.setAttribute("disabled", "");
            select.dataset.field = 'FORMDROPDOWN';
            select.classList.add(`${this.className}-field-formdropdown`);
            const items = ffData?.ddItems ?? [];
            const selectedText = children
                .map(n => (n instanceof Node ? (n.textContent ?? '') : String(n)))
                .join('')
                .trim();
            items.forEach((item, idx) => {
                const option = this.h({ tagName: "option" });
                option.setAttribute("value", item);
                option.textContent = item;
                if ((ffData?.ddDefault != null && ffData.ddDefault === idx) ||
                    (ffData?.ddDefault == null && item === selectedText)) {
                    option.setAttribute("selected", "");
                }
                select.appendChild(option);
            });
            if (items.length === 0 && selectedText) {
                const option = this.h({ tagName: "option" });
                option.setAttribute("value", selectedText);
                option.textContent = selectedText;
                option.setAttribute("selected", "");
                select.appendChild(option);
            }
            return [select];
        }
        wrapHyperlinkFieldResult(children, parsed) {
            const switchesLower = parsed.switches.map(s => s.toLowerCase());
            const hasLocal = switchesLower.includes('\\l');
            const hasNewWindow = switchesLower.includes('\\n');
            const tooltip = extractSwitchValue(parsed, '\\o');
            const targetFromT = extractSwitchValue(parsed, '\\t');
            const href = hasLocal
                ? '#' + (firstNonSwitchArg(parsed) ?? '')
                : (firstNonSwitchArg(parsed) ?? '');
            if (!hasLocal && !isSafeHyperlinkHref(href)) {
                const span = this.h({ tagName: "span" });
                children.forEach(c => span.appendChild(c));
                return [span];
            }
            const a = this.h({ tagName: "a" });
            a.setAttribute('href', href);
            children.forEach(c => a.appendChild(c));
            if (tooltip) {
                a.setAttribute('title', tooltip);
            }
            const effectiveTarget = hasNewWindow ? '_blank' : targetFromT;
            if (effectiveTarget && /^_(blank|self|parent|top)$/.test(effectiveTarget)) {
                a.setAttribute('target', effectiveTarget);
                if (!hasLocal) {
                    a.setAttribute('rel', 'noopener noreferrer');
                }
            }
            return [a];
        }
        renderContainer(elem, tagName) {
            return this.h({ tagName, children: this.renderElements(elem.children) });
        }
        renderContainerNS(elem, ns, tagName, props) {
            return this.h({ ns, tagName, children: this.renderElements(elem.children), ...props });
        }
        renderParagraph(elem) {
            this.applyParagraphBreakControls(elem);
            this.applyParagraphDropCap(elem);
            const tagName = getHeadingTagName(elem, this.styleMap);
            var result = this.toHTML(elem, ns.html, tagName);
            const style = this.findStyle(elem.styleName);
            elem.tabs ?? (elem.tabs = style?.paragraphProps?.tabs);
            const numbering = elem.numbering ?? style?.paragraphProps?.numbering;
            if (numbering) {
                result.classList.add(this.numberingClass(numbering.id, numbering.level));
            }
            if (this.showChanges && elem.paragraphMarkRevisionKind) {
                this.appendParagraphMarkRevision(result, elem);
            }
            this.applyFormattingRevision(result, elem);
            if (elem.paraId) {
                result.dataset.paraId = elem.paraId;
            }
            const isHeadingTag = /^H[1-6]$/.test(result.tagName);
            addSharedClass(result, isHeadingTag ? "heading" : "paragraph");
            if (!this.paragraphHasVisibleContent(result)) {
                result.appendChild(this.h({ tagName: "br" }));
            }
            return result;
        }
        paragraphHasVisibleContent(p) {
            const atomic = new Set(["BR", "IMG", "SVG", "MATH", "VIDEO", "CANVAS", "IFRAME", "OBJECT", "EMBED", "INPUT"]);
            const walk = (node) => {
                if (node.nodeType === 3) {
                    return (node.nodeValue ?? "").length > 0;
                }
                if (node.nodeType === 1) {
                    const el = node;
                    if (atomic.has(el.tagName))
                        return true;
                    for (const child of Array.from(el.childNodes)) {
                        if (walk(child))
                            return true;
                    }
                }
                return false;
            };
            for (const child of Array.from(p.childNodes)) {
                if (walk(child))
                    return true;
            }
            return false;
        }
        applyParagraphBreakControls(elem) {
            const css = elem.cssStyle ?? (elem.cssStyle = {});
            const style = this.findStyle(elem.styleName);
            const styleProps = style?.paragraphProps;
            const widowControl = elem.widowControl ?? styleProps?.widowControl;
            const keepNext = elem.keepNext ?? styleProps?.keepNext;
            const keepLines = elem.keepLines ?? styleProps?.keepLines;
            const pageBreakBefore = elem.pageBreakBefore ?? styleProps?.pageBreakBefore;
            if (widowControl === false) {
                css["widows"] = "0";
                css["orphans"] = "0";
            }
            if (keepNext === true && !css["break-after"]) {
                css["break-after"] = "avoid";
            }
            if (keepLines === true && !css["break-inside"]) {
                css["break-inside"] = "avoid";
            }
            if (pageBreakBefore === true && !css["break-before"]) {
                css["break-before"] = "page";
            }
        }
        applyParagraphDropCap(elem) {
            if (!elem.dropCap)
                return;
            const lines = (Number.isInteger(elem.dropCapLines) && elem.dropCapLines >= 1)
                ? elem.dropCapLines
                : 3;
            const css = elem.cssStyle ?? (elem.cssStyle = {});
            css["float"] = "left";
            css["font-size"] = `${lines}em`;
            css["line-height"] = "0.9";
            if (elem.dropCap === "drop") {
                css["margin"] = "0 0.1em 0 0";
            }
            else {
                css["margin-left"] = `-${lines * 0.5}em`;
            }
        }
        appendParagraphMarkRevision(paragraphEl, elem) {
            const c = this.className;
            const kind = elem.paragraphMarkRevisionKind;
            const rev = elem.revision;
            if (!kind)
                return;
            const classes = [`${c}-paragraph-mark`, `${c}-paragraph-mark-${kind}`];
            if (rev?.author && this.options.changes?.colorByAuthor !== false) {
                classes.push(`${c}-change-author-${this.getAuthorIndex(rev.author)}`);
            }
            const mark = this.h({
                tagName: "span",
                className: classes.join(" "),
                children: ["¶"]
            });
            if (rev?.id)
                mark.dataset.changeId = rev.id;
            if (rev?.author)
                mark.dataset.author = rev.author;
            if (rev?.date)
                mark.dataset.date = rev.date;
            mark.dataset.changeKind = "paragraphMark";
            mark.setAttribute("aria-label", kind === "inserted" ? "Paragraph inserted" : "Paragraph mark deleted");
            paragraphEl.appendChild(mark);
            this.changeElements.push(mark);
            this.changeMeta.push({
                el: mark, id: rev?.id, kind: "paragraphMark",
                author: rev?.author, date: rev?.date,
                summary: this.summarizeChange(mark, "paragraphMark"),
            });
        }
        renderHyperlink(elem) {
            const res = this.toH(elem, ns.html, "a");
            let rawHref = '';
            if (elem.id) {
                const rel = this.document.documentPart.rels.find(it => it.id == elem.id && it.targetMode === "External");
                rawHref = rel?.target ?? '';
            }
            if (rawHref && !isSafeHyperlinkHref(rawHref)) {
                return this.h({
                    ns: ns.html,
                    tagName: "span",
                    className: res.className,
                    style: res.style,
                    children: res.children,
                });
            }
            let href = rawHref;
            if (elem.anchor) {
                href += `#${elem.anchor}`;
            }
            res.href = href;
            const link = this.h(res);
            if (elem.tooltip) {
                link.setAttribute('title', elem.tooltip);
            }
            if (elem.targetFrame && /^_(blank|self|parent|top)$/.test(elem.targetFrame)) {
                link.setAttribute('target', elem.targetFrame);
                if (rawHref && !link.hasAttribute('rel')) {
                    link.setAttribute('rel', 'noopener noreferrer');
                }
            }
            return link;
        }
        renderSmartTag(elem) {
            return this.renderContainer(elem, "span");
        }
        renderRuby(elem) {
            const baseNodes = [];
            const rtNodes = [];
            for (const c of elem.children ?? []) {
                if (c.type === DomType.RubyText)
                    rtNodes.push(c);
                else
                    baseNodes.push(c);
            }
            const children = [
                ...this.renderElements(baseNodes),
                ...this.renderElements(rtNodes)
            ];
            const rubyEl = this.h({ tagName: "ruby", children });
            if (elem.rubyPr?.lid) {
                rubyEl.setAttribute("lang", elem.rubyPr.lid);
            }
            return rubyEl;
        }
        renderFitText(elem) {
            const widthTwips = typeof elem.width === "number" ? elem.width : parseFloat(elem.width);
            const children = this.renderElements(elem.children);
            if (!Number.isFinite(widthTwips) || widthTwips <= 0) {
                return this.h({ tagName: "span", children });
            }
            const pt = widthTwips / 20;
            const span = this.h({
                tagName: "span",
                children,
                style: {
                    "display": "inline-block",
                    "width": `${pt}pt`,
                    "white-space": "nowrap",
                    "overflow": "hidden"
                }
            });
            return span;
        }
        renderBidiOverride(elem) {
            const dir = elem.dir === "rtl" ? "rtl" : "ltr";
            const children = this.renderElements(elem.children);
            const bdo = this.h({ tagName: "bdo", children });
            bdo.setAttribute("dir", dir);
            return bdo;
        }
        renderSdt(elem) {
            const control = elem.sdtControl;
            const boundText = this.resolveSdtDataBinding(elem);
            let children;
            if (boundText != null) {
                const span = this.h({ tagName: "span" });
                span.textContent = boundText;
                span.dataset.sdtBound = "1";
                children = [span];
            }
            else if (control?.type === "checkbox") {
                const box = this.h({ tagName: "input" });
                box.setAttribute("type", "checkbox");
                box.setAttribute("disabled", "");
                if (control.checked)
                    box.setAttribute("checked", "");
                children = [box];
            }
            else if (control?.type === "dropdown") {
                const rendered = this.renderElements(elem.children) ?? [];
                const selectedText = rendered
                    .map(n => (n instanceof Node ? (n.textContent ?? "") : String(n)))
                    .join("")
                    .trim();
                const select = this.h({ tagName: "select" });
                select.setAttribute("disabled", "");
                for (const item of control.items) {
                    const option = this.h({ tagName: "option" });
                    option.setAttribute("value", item.value);
                    option.textContent = item.displayText;
                    if (item.value === selectedText ||
                        item.displayText === selectedText) {
                        option.setAttribute("selected", "");
                    }
                    select.appendChild(option);
                }
                children = [select];
            }
            else if (control?.type === "date") {
                const rendered = this.renderElements(elem.children) ?? [];
                const time = this.h({ tagName: "time" });
                const ISO = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2}(\.\d+)?(Z|[+-]\d{2}:\d{2})?)?$/;
                if (control.fullDate && ISO.test(control.fullDate)) {
                    time.setAttribute("datetime", control.fullDate);
                }
                for (const c of rendered)
                    time.appendChild(c);
                children = [time];
            }
            else {
                children = this.renderElements(elem.children) ?? [];
            }
            const span = this.h({ tagName: "span", children });
            span.setAttribute("role", "group");
            if (elem.sdtAlias) {
                span.setAttribute("aria-label", elem.sdtAlias);
            }
            if (elem.sdtTag) {
                span.dataset.sdtTag = elem.sdtTag;
            }
            if (control) {
                span.dataset.sdtType = control.type;
            }
            return span;
        }
        resolveSdtDataBinding(elem) {
            const binding = elem.dataBinding;
            if (!binding?.xpath)
                return null;
            if (!this.document)
                return null;
            if (!isSafeCustomXmlXPath(binding.xpath))
                return null;
            const storeItemID = binding.storeItemID;
            const parts = [];
            if (storeItemID) {
                const part = this.document.findCustomXmlByStoreItemId?.(storeItemID);
                if (part)
                    parts.push(part);
            }
            else {
                for (const p of this.document.customXmlParts ?? []) {
                    if (p.xmlDoc)
                        parts.push(p);
                }
            }
            if (parts.length === 0)
                return null;
            const evalDoc = typeof document !== 'undefined' ? document : null;
            for (const part of parts) {
                if (!part.xmlDoc)
                    continue;
                const result = safeEvaluateXPath(evalDoc, part.xmlDoc, binding.xpath);
                if (result != null && result !== '')
                    return result;
            }
            return null;
        }
        renderCommentRangeStart(commentStart) {
            if (!this.options.renderComments) {
                const anchor = this.h({
                    tagName: "span",
                    className: `${this.className}-comment-anchor-start comment-reference`,
                });
                if (commentStart.id)
                    anchor.dataset.comment = commentStart.id;
                return anchor;
            }
            if (this.useSidebar) {
                const anchor = this.h({ tagName: "span", className: `${this.className}-comment-anchor-start` });
                anchor.dataset.commentId = commentStart.id;
                anchor.dataset.comment = '';
                if (!this.commentAnchorElements[commentStart.id]) {
                    this.commentAnchorElements[commentStart.id] = [];
                }
                this.commentAnchorElements[commentStart.id].push(anchor);
                if (this.useHighlight) {
                    const rng = new Range();
                    this.commentHighlight?.add(rng);
                    this.later(() => rng.setStart(anchor, 0));
                    this.commentMap[commentStart.id] = rng;
                }
                this.later(() => {
                    anchor.addEventListener("click", () => {
                        const sidebarEl = this.sidebarCommentElements[commentStart.id];
                        if (sidebarEl) {
                            sidebarEl.scrollIntoView({ behavior: "smooth", block: "center" });
                            sidebarEl.classList.add(`${this.className}-comment-focused`);
                            setTimeout(() => sidebarEl.classList.remove(`${this.className}-comment-focused`), 2000);
                        }
                    });
                });
                return anchor;
            }
            const rng = new Range();
            this.commentHighlight?.add(rng);
            const result = this.h({ tagName: "#comment", children: [`start of comment #${commentStart.id}`] });
            this.later(() => rng.setStart(result, 0));
            this.commentMap[commentStart.id] = rng;
            return result;
        }
        renderCommentRangeEnd(commentEnd) {
            if (!this.options.renderComments)
                return null;
            if (this.useSidebar) {
                const anchor = this.h({ tagName: "span", className: `${this.className}-comment-anchor-end` });
                anchor.dataset.commentId = commentEnd.id;
                anchor.dataset.comment = '';
                if (this.useHighlight) {
                    const rng = this.commentMap[commentEnd.id];
                    this.later(() => rng?.setEnd(anchor, 0));
                }
                return anchor;
            }
            const rng = this.commentMap[commentEnd.id];
            const result = this.h({ tagName: "#comment", children: [`end of comment #${commentEnd.id}`] });
            this.later(() => rng?.setEnd(result, 0));
            return result;
        }
        renderCommentReference(commentRef) {
            if (!this.options.renderComments) {
                const anchor = this.h({
                    tagName: "sup",
                    className: `${this.className}-comment-ref comment-reference`,
                });
                if (commentRef.id)
                    anchor.dataset.comment = commentRef.id;
                return anchor;
            }
            if (this.useSidebar) {
                return this.h({ tagName: "#comment", children: [`comment ref #${commentRef.id}`] });
            }
            var comment = this.document.commentsPart?.commentMap[commentRef.id];
            if (!comment)
                return null;
            const commentRefEl = this.h({ tagName: "span", className: `${this.className}-comment-ref comment-reference`, children: ['💬'] });
            if (commentRef.id)
                commentRefEl.dataset.comment = commentRef.id;
            const commentsContainerEl = this.h({
                tagName: "div", className: `${this.className}-comment-popover`, children: [
                    this.h({ tagName: 'div', className: `${this.className}-comment-author`, children: [comment.author] }),
                    this.h({ tagName: 'div', className: `${this.className}-comment-date`, children: [new Date(comment.date).toLocaleString()] }),
                    ...this.renderElements(comment.children)
                ]
            });
            return this.h({ tagName: "#fragment", children: [
                    this.h({ tagName: "#comment", children: [`comment #${comment.id} by ${comment.author} on ${comment.date}`] }),
                    commentRefEl,
                    commentsContainerEl
                ] });
        }
        renderDrawing(elem) {
            var result = this.toHTML(elem, ns.html, "div");
            const parsed = elem.cssStyle ?? {};
            if (!parsed["display"] && !parsed["float"])
                result.style.display = "inline-block";
            if (!parsed["position"] && !parsed["float"])
                result.style.position = "relative";
            result.style.textIndent = "0px";
            if (elem.props?.isAnchor) {
                result.setAttribute("data-drawing-anchor", "true");
            }
            return result;
        }
        emuToPx(emu) {
            return (emu ?? 0) / 9525;
        }
        buildShapeRenderContext() {
            const self = this;
            const paletteSrc = this.document?.themePart?.theme?.colorScheme?.colors ?? null;
            return {
                nextId(prefix) {
                    self._shapeIdCounter += 1;
                    return `${prefix}-${self._shapeIdCounter}`;
                },
                themePalette: paletteSrc,
            };
        }
        renderDrawingShape(elem) {
            const result = renderShape(elem, (emu) => this.emuToPx(emu), (paragraphs) => this.renderElements(paragraphs), this.buildShapeRenderContext());
            return result;
        }
        renderDrawingShapeGroup(elem) {
            return renderShapeGroup(elem, (emu) => this.emuToPx(emu), (child) => {
                const rendered = this.renderElement(child);
                if (!rendered)
                    return null;
                if (Array.isArray(rendered))
                    return rendered[0] ?? null;
                return rendered;
            }, this.buildShapeRenderContext());
        }
        renderImage(elem) {
            let result = this.toHTML(elem, ns.html, "img", []);
            result.alt = elem.altText ?? "";
            let transform = elem.cssStyle?.transform;
            if (elem.srcRect && elem.srcRect.some(x => x != 0)) {
                var [left, top, right, bottom] = elem.srcRect;
                transform = `scale(${1 / (1 - left - right)}, ${1 / (1 - top - bottom)})`;
                result.style['clip-path'] = `rect(${(100 * top).toFixed(2)}% ${(100 * (1 - right)).toFixed(2)}% ${(100 * (1 - bottom)).toFixed(2)}% ${(100 * left).toFixed(2)}%)`;
            }
            if (elem.rotation)
                transform = `rotate(${elem.rotation}deg) ${transform ?? ''}`;
            result.style.transform = transform?.trim();
            if (this.document) {
                this.tasks.push(this.document.loadDocumentImage(elem.src, this.currentPart).then(x => {
                    result.src = x;
                }));
            }
            addSharedClass(result, "image");
            return result;
        }
        renderChart(elem) {
            const fallback = this.createElement("span");
            fallback.className = `${this.className}-chart`;
            fallback.style.display = "inline-block";
            if (!elem.relId || !this.document)
                return fallback;
            const part = this.document.findPartByRelId(elem.relId, this.currentPart ?? this.document.documentPart);
            if (!part || !(part instanceof ChartPart) || !part.chart)
                return fallback;
            try {
                const themePalette = this.document?.themePart?.theme?.colorScheme?.colors ?? null;
                const svg = renderChart(part.chart, { themePalette });
                const wrapper = this.createElement("span");
                wrapper.className = `${this.className}-chart`;
                wrapper.style.display = "inline-block";
                wrapper.appendChild(svg);
                scheduleLegendOverflowAdjust(svg);
                return wrapper;
            }
            catch {
                return fallback;
            }
        }
        renderChartEx(elem) {
            const fallback = this.createElement("div");
            fallback.className = `${this.className}-chartex-placeholder`;
            if (!elem.relId || !this.document)
                return fallback;
            const part = this.document.findPartByRelId(elem.relId, this.currentPart ?? this.document.documentPart);
            if (!part || !(part instanceof ChartExPart) || !part.chart)
                return fallback;
            const chart = part.chart;
            if (chart.shape === "data") {
                try {
                    let svg = null;
                    switch (chart.kind) {
                        case "sunburst":
                            svg = renderSunburst(chart);
                            break;
                        case "treemap":
                            svg = renderTreemap(chart);
                            break;
                        case "waterfall":
                            svg = renderWaterfall(chart);
                            break;
                        case "funnel":
                            svg = renderFunnel(chart);
                            break;
                        case "histogram":
                            svg = renderHistogram(chart);
                            break;
                    }
                    if (svg) {
                        const wrapper = this.createElement("span");
                        wrapper.className = `${this.className}-chart`;
                        wrapper.style.display = "inline-block";
                        wrapper.setAttribute("data-chart-kind", chart.kind);
                        wrapper.appendChild(svg);
                        return wrapper;
                    }
                }
                catch {
                }
            }
            const wrapper = this.createElement("div");
            wrapper.className = `docx-chartex-placeholder`;
            wrapper.setAttribute("data-chart-kind", chart.kind);
            const titleDiv = this.createElement("div");
            titleDiv.className = "docx-chartex-placeholder__title";
            titleDiv.textContent = chart.title || "";
            wrapper.appendChild(titleDiv);
            const noteDiv = this.createElement("div");
            noteDiv.className = "docx-chartex-placeholder__note";
            noteDiv.textContent = "Chart type not yet supported";
            wrapper.appendChild(noteDiv);
            return wrapper;
        }
        renderSmartArtPlaceholder(elem) {
            const wrapper = this.createElement("div");
            wrapper.className = "docx-smartart-placeholder";
            let layoutId = elem.layoutId ?? "";
            const loRelId = elem.relIds?.lo;
            if (!layoutId && loRelId && this.document) {
                const part = this.document.findPartByRelId(loRelId, this.currentPart ?? this.document.documentPart);
                if (part instanceof DiagramLayoutPart && part.layoutId) {
                    layoutId = part.layoutId;
                }
            }
            if (layoutId) {
                wrapper.setAttribute("data-smartart-layout", layoutId);
            }
            const noteDiv = this.createElement("div");
            noteDiv.className = "docx-smartart-placeholder__note";
            noteDiv.textContent = "SmartArt diagram not yet supported";
            wrapper.appendChild(noteDiv);
            return wrapper;
        }
        renderText(elem) {
            return this.h(elem.text);
        }
        renderDeletedText(elem) {
            return (this.showChanges && this.options.changes?.showDeletions !== false)
                ? this.renderText(elem)
                : null;
        }
        renderBreak(elem) {
            if (elem.break == "textWrapping") {
                return this.h({ tagName: "br" });
            }
            if (elem.break == "page") {
                const br = this.h({ tagName: "br" });
                br.classList.add(`${this.className}-page-break`);
                br.classList.add("page-break");
                br.dataset.pageBreak = "";
                return br;
            }
            return null;
        }
        renderInserted(elem) {
            if (this.showChanges && this.options.changes?.showInsertions !== false) {
                const node = this.renderContainer(elem, "ins");
                this.applyChangeAttributes(node, elem, "insertion");
                return node;
            }
            return this.renderElements(elem.children);
        }
        renderDeleted(elem) {
            if (this.showChanges && this.options.changes?.showDeletions !== false) {
                const node = this.renderContainer(elem, "del");
                this.applyChangeAttributes(node, elem, "deletion");
                return node;
            }
            return null;
        }
        renderMoveFrom(elem) {
            if (!this.showChanges || this.options.changes?.showMoves === false) {
                return null;
            }
            const node = this.renderContainer(elem, "span");
            node.classList.add(`${this.className}-move-from`);
            this.applyChangeAttributes(node, elem, "move");
            this.registerMove(node, elem, "from");
            return node;
        }
        renderMoveTo(elem) {
            if (!this.showChanges || this.options.changes?.showMoves === false) {
                return this.renderElements(elem.children);
            }
            const node = this.renderContainer(elem, "span");
            node.classList.add(`${this.className}-move-to`);
            this.applyChangeAttributes(node, elem, "move");
            this.registerMove(node, elem, "to");
            return node;
        }
        registerMove(node, elem, half) {
            const id = elem.revision?.id;
            if (!id)
                return;
            node.dataset.moveId = id;
            const pair = this.moveElements.get(id) ?? {};
            pair[half] = node;
            this.moveElements.set(id, pair);
            this.later(() => {
                node.addEventListener("click", (ev) => {
                    const entry = this.moveElements.get(id);
                    const counterpart = half === "from" ? entry?.to : entry?.from;
                    if (counterpart) {
                        ev.preventDefault();
                        counterpart.scrollIntoView({ behavior: "smooth", block: "center" });
                    }
                });
            });
        }
        applyChangeAttributes(node, elem, kind) {
            const rev = elem.revision;
            if (!rev)
                return;
            if (rev.id)
                node.dataset.changeId = rev.id;
            if (rev.author)
                node.dataset.author = rev.author;
            if (rev.date)
                node.dataset.date = rev.date;
            node.dataset.changeKind = kind;
            if (rev.author && this.options.changes?.colorByAuthor !== false) {
                const idx = this.getAuthorIndex(rev.author);
                node.classList.add(`${this.className}-change-author-${idx}`);
            }
            this.changeElements.push(node);
            this.changeMeta.push({
                el: node,
                id: rev.id,
                kind,
                author: rev.author,
                date: rev.date,
                summary: this.summarizeChange(node, kind),
            });
        }
        summarizeChange(node, kind) {
            const MAX = 80;
            const truncate = (s) => {
                const clean = s.replace(/\s+/g, " ").trim();
                return clean.length > MAX ? clean.slice(0, MAX - 1) + "…" : clean;
            };
            switch (kind) {
                case "insertion":
                case "move": {
                    const text = truncate(node.textContent ?? "");
                    return text ? `Inserted: "${text}"` : "Inserted content";
                }
                case "deletion": {
                    const text = truncate(node.textContent ?? "");
                    return text ? `Deleted: "${text}"` : "Deleted content";
                }
                case "paragraphMark":
                    return "Paragraph mark changed";
                case "rowInsertion":
                    return "Row inserted";
                case "rowDeletion":
                    return "Row deleted";
                case "formatting": {
                    const title = node.getAttribute("title");
                    return title ?? "Formatting changed";
                }
            }
        }
        getAuthorIndex(author) {
            let idx = this.changeAuthorIndex.get(author);
            if (idx === undefined) {
                idx = this.changeAuthorIndex.size % HtmlRenderer.CHANGE_PALETTE_SIZE;
                this.changeAuthorIndex.set(author, idx);
            }
            return idx;
        }
        renderSymbol(elem) {
            return this.h({ tagName: "span", children: [String.fromCharCode(elem.char)], style: { fontFamily: elem.font } });
        }
        renderFootnoteReference(elem) {
            this.currentFootnoteIds.push(elem.id);
            this.footnoteRefCount++;
            const sup = this.h({
                tagName: "sup",
                className: `${this.className}-footnote-ref footnote-ref footnote`,
                children: [`${this.footnoteRefCount}`]
            });
            if (elem.id)
                sup.dataset.footnoteId = elem.id;
            sup.dataset.footnote = '';
            return sup;
        }
        renderEndnoteReference(elem) {
            this.currentEndnoteIds.push(elem.id);
            this.endnoteRefCount++;
            const sup = this.h({
                tagName: "sup",
                className: `${this.className}-endnote-ref`,
                children: [`${this.endnoteRefCount}`]
            });
            if (elem.id)
                sup.dataset.footnoteId = elem.id;
            return sup;
        }
        renderTab(elem) {
            var tabSpan = this.h({ tagName: "span", children: ["\u2003\u2003\u2003\u2003"] });
            if (this.options.experimental) {
                tabSpan.className = this.tabStopClass();
                var stops = findParent(elem, DomType.Paragraph)?.tabs;
                this.currentTabs.push({ stops, span: tabSpan });
            }
            return tabSpan;
        }
        renderBookmarkStart(elem) {
            return this.h({ tagName: "span", id: elem.name });
        }
        renderRun(elem, bypassFieldGuard = false) {
            if (elem.fieldRun && !bypassFieldGuard)
                return null;
            let children = this.renderElements(elem.children);
            if (elem.verticalAlign === "sub" || elem.verticalAlign === "sup") {
                children = [this.h({ tagName: elem.verticalAlign, children })];
            }
            const result = this.toHTML(elem, ns.html, "span", children);
            if (elem.id)
                result.id = elem.id;
            this.applyFormattingRevision(result, elem);
            addSharedClass(result, "run");
            return result;
        }
        applyFormattingRevision(node, elem) {
            const fr = elem.formattingRevision;
            if (!fr)
                return;
            if (!this.showChanges || this.options.changes?.showFormatting === false)
                return;
            const c = this.className;
            node.classList.add(`${c}-formatting-revision`);
            if (fr.id)
                node.dataset.changeId = fr.id;
            if (fr.author)
                node.dataset.author = fr.author;
            if (fr.date)
                node.dataset.date = fr.date;
            if (fr.author && this.options.changes?.colorByAuthor !== false) {
                node.classList.add(`${c}-change-author-${this.getAuthorIndex(fr.author)}`);
            }
            const changed = fr.changedProps && fr.changedProps.length
                ? fr.changedProps.join(", ")
                : "formatting";
            const who = fr.author ? `${fr.author} changed` : "Changed";
            node.setAttribute("title", `${who}: ${changed}`);
            node.dataset.changeKind = "formatting";
            if (fr.id)
                node.dataset.changeId = fr.id;
            this.changeElements.push(node);
            this.changeMeta.push({
                el: node, id: fr.id, kind: "formatting",
                author: fr.author, date: fr.date,
                summary: `${who}: ${changed}`,
            });
        }
        availableContentWidthPt() {
            const sect = this.currentSectProps;
            if (!sect)
                return null;
            const pageW = sect.pageSize?.width;
            const left = sect.pageMargins?.left;
            const right = sect.pageMargins?.right;
            if (!pageW || left == null || right == null)
                return null;
            if (!/pt\s*$/.test(pageW) || !/pt\s*$/.test(left) || !/pt\s*$/.test(right))
                return null;
            const w = parseFloat(pageW);
            const l = parseFloat(left);
            const r = parseFloat(right);
            if (!Number.isFinite(w) || !Number.isFinite(l) || !Number.isFinite(r))
                return null;
            const avail = w - l - r;
            return avail > 0 ? avail : null;
        }
        renderTable(elem) {
            const isNested = this.tableCellPositions.length > 0;
            this.tableCellPositions.push(this.currentCellPosition);
            this.tableVerticalMerges.push(this.currentVerticalMerge);
            this.currentVerticalMerge = {};
            this.currentCellPosition = { col: 0, row: 0 };
            const availablePt = this.availableContentWidthPt();
            if (elem.cssStyle && availablePt != null) {
                const widthStr = elem.cssStyle["width"];
                if (widthStr && /pt\s*$/.test(widthStr)) {
                    const widthPt = parseFloat(widthStr);
                    if (Number.isFinite(widthPt) && widthPt > availablePt) {
                        elem.cssStyle["width"] = `${availablePt.toFixed(2)}pt`;
                    }
                }
                if (elem.cssStyle["max-width"] == null) {
                    elem.cssStyle["max-width"] = "100%";
                }
            }
            else if (elem.cssStyle && elem.cssStyle["max-width"] == null) {
                elem.cssStyle["max-width"] = "100%";
            }
            this.tableBandSizes.push(this.currentTableBandSizes);
            this.currentTableBandSizes = {
                col: Math.max(1, elem.colBandSize ?? 1),
                row: Math.max(1, elem.rowBandSize ?? 1),
            };
            const children = [];
            if (elem.columns)
                children.push(this.renderTableColumns(elem.columns));
            const headerRendered = [];
            const bodyRendered = [];
            const anyRowExplicitHeader = (elem.children ?? []).some((c) => c.type === DomType.Row
                && c.isHeader !== undefined
                && c.isHeader !== false);
            const promoteFirstRow = !!elem.firstRowIsHeader && !anyRowExplicitHeader;
            let firstRowPromoted = false;
            for (const child of (elem.children ?? [])) {
                const rendered = this.renderElement(child);
                if (rendered == null)
                    continue;
                let isHeaderRow = child.type === DomType.Row
                    && child.isHeader !== undefined
                    && child.isHeader !== false;
                if (!isHeaderRow && promoteFirstRow && !firstRowPromoted
                    && child.type === DomType.Row) {
                    isHeaderRow = true;
                    firstRowPromoted = true;
                }
                if (isHeaderRow) {
                    if (Array.isArray(rendered)) {
                        for (const node of rendered)
                            if (node instanceof HTMLElement)
                                headerRendered.push(node);
                    }
                    else if (rendered instanceof HTMLElement) {
                        headerRendered.push(rendered);
                    }
                }
                else {
                    if (Array.isArray(rendered))
                        bodyRendered.push(...rendered);
                    else
                        bodyRendered.push(rendered);
                }
            }
            if (headerRendered.length > 0) {
                for (const tr of headerRendered) {
                    if (tr.tagName === "TR")
                        tr.setAttribute("data-header", "true");
                }
                children.push(this.h({ tagName: "thead", children: headerRendered }));
            }
            if (bodyRendered.length > 0) {
                children.push(this.h({ tagName: "tbody", children: bodyRendered }));
            }
            this.currentVerticalMerge = this.tableVerticalMerges.pop();
            this.currentCellPosition = this.tableCellPositions.pop();
            this.currentTableBandSizes = this.tableBandSizes.pop();
            if (isNested && elem.cssStyle) {
                const inherited = elem.cssStyle;
                if (inherited["background-color"] == null)
                    inherited["background-color"] = "transparent";
                if (inherited["background-image"] == null)
                    inherited["background-image"] = "none";
            }
            const tableResult = this.toHTML(elem, ns.html, "table", children);
            addSharedClass(tableResult, "table");
            this.applyFirstRowHeaderA11y(tableResult, elem);
            return tableResult;
        }
        applyFirstRowHeaderA11y(table, elem) {
            const hasFirstRowStyle = elem.className?.split(/\s+/).includes("first-row");
            if (!hasFirstRowStyle)
                return;
            const thead = table.querySelector(":scope > thead");
            if (thead && thead.querySelector("th, td"))
                return;
            const tbody = table.querySelector(":scope > tbody") ?? table;
            const firstRow = tbody.querySelector(":scope > tr");
            if (!firstRow)
                return;
            const cells = Array.from(firstRow.querySelectorAll(":scope > td, :scope > th"))
                .filter(c => c.style.display !== "none");
            if (cells.length === 0)
                return;
            const gridCols = table.querySelectorAll(":scope > colgroup > col").length
                || cells.reduce((n, c) => n + (c.colSpan || 1), 0);
            if (cells.length === 1) {
                const only = cells[0];
                if ((only.colSpan || 1) >= gridCols) {
                    only.setAttribute("scope", "colgroup");
                }
                else {
                    only.setAttribute("scope", "col");
                }
                return;
            }
            for (const cell of cells) {
                cell.setAttribute("scope", "col");
            }
        }
        renderTableColumns(columns) {
            const children = columns.map(x => this.h({ tagName: "col", style: { width: x.width } }));
            return this.h({ tagName: "colgroup", children });
        }
        renderTableRow(elem) {
            this.currentCellPosition.col = 0;
            const children = [];
            if (elem.gridBefore)
                children.push(this.renderTableCellPlaceholder(elem.gridBefore));
            const rowBookmarks = [];
            const cellChildren = [];
            for (const child of (elem.children ?? [])) {
                if (child.type === DomType.BookmarkStart) {
                    const bm = child;
                    if (bm.colFirst != null && bm.colLast != null && bm.name) {
                        rowBookmarks.push({ name: bm.name, colFirst: bm.colFirst, colLast: bm.colLast });
                        continue;
                    }
                }
                if (child.type === DomType.BookmarkEnd)
                    continue;
                cellChildren.push(child);
            }
            const prevHeader = this.currentRowIsHeader;
            this.currentRowIsHeader = elem.isHeader !== undefined && elem.isHeader !== false;
            const renderedCells = this.renderElements(cellChildren);
            this.currentRowIsHeader = prevHeader;
            children.push(...renderedCells);
            if (elem.gridAfter)
                children.push(this.renderTableCellPlaceholder(elem.gridAfter));
            this.currentCellPosition.row++;
            const tr = this.toHTML(elem, ns.html, "tr", children);
            if (elem.cantSplit) {
                tr.setAttribute("data-cant-split", "");
            }
            const rowBandSize = this.currentTableBandSizes.row;
            if (rowBandSize > 0) {
                const rowIdx = Math.max(0, (this.currentCellPosition.row - 1) | 0);
                tr.setAttribute("data-band", String(Math.floor(rowIdx / rowBandSize) % 2));
            }
            if (rowBookmarks.length > 0) {
                const cellNodes = [];
                let idx = 0;
                for (const child of cellChildren) {
                    if (child.type === DomType.Cell) {
                        const node = renderedCells[idx];
                        if (node instanceof HTMLElement)
                            cellNodes.push(node);
                    }
                    idx++;
                }
                const ranges = [];
                for (const bm of rowBookmarks) {
                    ranges.push(`${bm.colFirst}-${bm.colLast}`);
                    const targetCell = cellNodes[bm.colFirst];
                    if (!targetCell)
                        continue;
                    const anchor = document.createElement('span');
                    anchor.setAttribute('id', bm.name);
                    targetCell.insertBefore(anchor, targetCell.firstChild);
                }
                tr.setAttribute('data-bookmark-cols', ranges.join(','));
            }
            if (this.showChanges && elem.rowRevisionKind) {
                this.applyRowRevision(tr, elem);
            }
            this.applyFormattingRevision(tr, elem);
            addSharedClass(tr, "table-row");
            return tr;
        }
        applyRowRevision(tr, elem) {
            const kind = elem.rowRevisionKind;
            if (!kind)
                return;
            if (kind === "inserted" && this.options.changes?.showInsertions === false)
                return;
            if (kind === "deleted" && this.options.changes?.showDeletions === false)
                return;
            const c = this.className;
            tr.classList.add(`${c}-row-${kind}`);
            const rev = elem.revision;
            if (rev?.id)
                tr.dataset.changeId = rev.id;
            if (rev?.author)
                tr.dataset.author = rev.author;
            if (rev?.date)
                tr.dataset.date = rev.date;
            const metaKind = kind === "inserted" ? "rowInsertion" : "rowDeletion";
            tr.dataset.changeKind = metaKind;
            if (rev?.author && this.options.changes?.colorByAuthor !== false) {
                tr.classList.add(`${c}-change-author-${this.getAuthorIndex(rev.author)}`);
            }
            this.changeElements.push(tr);
            this.changeMeta.push({
                el: tr, id: rev?.id, kind: metaKind,
                author: rev?.author, date: rev?.date,
                summary: this.summarizeChange(tr, metaKind),
            });
        }
        renderTableCellPlaceholder(colSpan) {
            return this.h({ tagName: "td", colSpan, style: { border: "none" } });
        }
        renderTableCell(elem) {
            const tagName = this.currentRowIsHeader ? "th" : "td";
            const diagTlBr = elem.cssStyle?.["$diag-tlbr"];
            const diagTrBl = elem.cssStyle?.["$diag-trbl"];
            let result = this.toHTML(elem, ns.html, tagName);
            if (this.currentRowIsHeader) {
                result.setAttribute("scope", "col");
            }
            const key = this.currentCellPosition.col;
            if (elem.verticalMerge) {
                if (elem.verticalMerge == "restart") {
                    this.currentVerticalMerge[key] = result;
                    result.rowSpan = 1;
                }
                else if (this.currentVerticalMerge[key]) {
                    this.currentVerticalMerge[key].rowSpan += 1;
                    result.style.display = "none";
                }
            }
            else {
                this.currentVerticalMerge[key] = null;
            }
            if (elem.span)
                result.colSpan = elem.span;
            const colBandSize = this.currentTableBandSizes.col;
            if (colBandSize > 0) {
                result.setAttribute("data-band", String(Math.floor(this.currentCellPosition.col / colBandSize) % 2));
            }
            this.currentCellPosition.col += result.colSpan;
            if (diagTlBr || diagTrBl) {
                this.applyDiagonalBorders(result, diagTlBr, diagTrBl);
            }
            addSharedClass(result, "table-cell");
            return result;
        }
        parseBorderStroke(border) {
            if (!border || border === "none")
                return { color: "black", width: "1" };
            const parts = border.split(/\s+/);
            if (parts.length < 3)
                return { color: "black", width: "1" };
            const w = parts[0].replace(/pt$/, "");
            const color = parts.slice(2).join(" ");
            return { color, width: w || "1" };
        }
        applyDiagonalBorders(cell, tlBr, trBl) {
            const existing = Array.from(cell.childNodes);
            const content = this.h({ tagName: "div", style: { position: "relative", zIndex: "1" } });
            for (const node of existing)
                content.appendChild(node);
            const overlay = this.h({
                ns: ns.svg,
                tagName: "svg",
                style: { position: "absolute", inset: "0", width: "100%", height: "100%", pointerEvents: "none", zIndex: "0" },
                preserveAspectRatio: "none",
                viewBox: "0 0 100 100",
            });
            if (tlBr) {
                const { color, width } = this.parseBorderStroke(tlBr);
                const line = this.h({
                    ns: ns.svg,
                    tagName: "line",
                    x1: "0", y1: "0", x2: "100", y2: "100",
                    stroke: color,
                    "stroke-width": width,
                    vectorEffect: "non-scaling-stroke",
                });
                overlay.appendChild(line);
            }
            if (trBl) {
                const { color, width } = this.parseBorderStroke(trBl);
                const line = this.h({
                    ns: ns.svg,
                    tagName: "line",
                    x1: "100", y1: "0", x2: "0", y2: "100",
                    stroke: color,
                    "stroke-width": width,
                    vectorEffect: "non-scaling-stroke",
                });
                overlay.appendChild(line);
            }
            cell.style.position = cell.style.position || "relative";
            cell.appendChild(overlay);
            cell.appendChild(content);
        }
        renderVmlPicture(elem) {
            return this.renderContainer(elem, "div");
        }
        renderOleObject(elem) {
            const progId = elem.progId;
            const allowed = progId && Object.prototype.hasOwnProperty.call(HtmlRenderer.OLE_PROGID_LABELS, progId);
            const label = allowed ? HtmlRenderer.OLE_PROGID_LABELS[progId] : 'Embedded object';
            const span = this.h({
                tagName: "span",
                className: `${this.className}-ole-placeholder`,
                children: [label],
            });
            span.dataset.progid = allowed ? progId : '';
            span.setAttribute("title", `Embedded object: ${label}`);
            span.setAttribute("role", "img");
            span.setAttribute("aria-label", label);
            return span;
        }
        renderVmlElement(elem) {
            var container = this.h({ ns: ns.svg, tagName: "svg", style: elem.cssStyleText });
            const result = this.renderVmlChildElement(elem);
            if (elem.imageHref?.id) {
                this.tasks.push(this.document?.loadDocumentImage(elem.imageHref.id, this.currentPart)
                    .then(x => result.setAttribute("href", x)));
            }
            container.appendChild(result);
            requestAnimationFrame(() => {
                const bb = container.firstElementChild.getBBox();
                const w = Math.max(1, Math.ceil(bb.width));
                const h = Math.max(1, Math.ceil(bb.height));
                container.setAttribute("width", `${w}`);
                container.setAttribute("height", `${h}`);
                container.setAttribute("viewBox", `${Math.floor(bb.x)} ${Math.floor(bb.y)} ${w} ${h}`);
            });
            return container;
        }
        renderVmlChildElement(elem) {
            const result = this.createSvgElement(elem.tagName);
            Object.entries(elem.attrs).forEach(([k, v]) => result.setAttribute(k, v));
            for (let child of elem.children) {
                if (child.type == DomType.VmlElement) {
                    result.appendChild(this.renderVmlChildElement(child));
                }
                else {
                    result.appendChild(...asArray(this.renderElement(child)));
                }
            }
            return result;
        }
        renderMmlRadical(elem) {
            const base = elem.children.find(el => el.type == DomType.MmlBase);
            if (elem.props?.hideDegree) {
                return this.createMathMLElement("msqrt", null, this.renderElements([base]));
            }
            const degree = elem.children.find(el => el.type == DomType.MmlDegree);
            return this.createMathMLElement("mroot", null, this.renderElements([base, degree]));
        }
        renderMmlDelimiter(elem) {
            const children = [];
            children.push(this.createMathMLElement("mo", null, [elem.props.beginChar ?? '(']));
            children.push(...this.renderElements(elem.children));
            children.push(this.createMathMLElement("mo", null, [elem.props.endChar ?? ')']));
            return this.createMathMLElement("mrow", null, children);
        }
        renderMmlNary(elem) {
            const children = [];
            const grouped = keyBy(elem.children, x => x.type);
            const sup = grouped[DomType.MmlSuperArgument];
            const sub = grouped[DomType.MmlSubArgument];
            const supElem = sup ? this.createMathMLElement("mo", null, asArray(this.renderElement(sup))) : null;
            const subElem = sub ? this.createMathMLElement("mo", null, asArray(this.renderElement(sub))) : null;
            const opChar = elem.props?.char ?? '\u222B';
            const charElem = this.createMathMLElement("mo", null, [opChar]);
            if (supElem || subElem) {
                if (supElem && subElem) {
                    const tag = resolveNaryLimitTag(elem.props?.limLoc, opChar);
                    children.push(this.createMathMLElement(tag, null, [charElem, subElem, supElem]));
                }
                else if (supElem) {
                    const tag = resolveNaryLimitTag(elem.props?.limLoc, opChar) === "munderover" ? "mover" : "msup";
                    children.push(this.createMathMLElement(tag, null, [charElem, supElem]));
                }
                else {
                    const tag = resolveNaryLimitTag(elem.props?.limLoc, opChar) === "munderover" ? "munder" : "msub";
                    children.push(this.createMathMLElement(tag, null, [charElem, subElem]));
                }
            }
            else {
                children.push(charElem);
            }
            children.push(...this.renderElements(grouped[DomType.MmlBase].children));
            return this.createMathMLElement("mrow", null, children);
        }
        renderMmlAccent(elem) {
            const base = elem.children.find(el => el.type == DomType.MmlBase);
            const baseNodes = base ? asArray(this.renderElement(base)) : [];
            const baseElem = this.createMathMLElement("mrow", null, baseNodes);
            const accentChar = elem.props?.char ?? '\u00AF';
            const accentElem = this.createMathMLElement("mo", null, [accentChar]);
            return this.createMathMLElement("mover", null, [baseElem, accentElem]);
        }
        renderMmlBorderBox(elem) {
            const result = this.createMathMLElement("menclose", null, this.renderElements(elem.children));
            result.setAttribute("notation", "box");
            return result;
        }
        renderMmlSubSuperscript(elem) {
            const grouped = keyBy(elem.children, x => x.type);
            const base = grouped[DomType.MmlBase];
            const sub = grouped[DomType.MmlSubArgument];
            const sup = grouped[DomType.MmlSuperArgument];
            const baseElem = base
                ? this.createMathMLElement("mrow", null, asArray(this.renderElement(base)))
                : this.createMathMLElement("mrow", null, []);
            const subElem = sub
                ? this.createMathMLElement("mrow", null, asArray(this.renderElement(sub)))
                : this.createMathMLElement("mrow", null, []);
            const supElem = sup
                ? this.createMathMLElement("mrow", null, asArray(this.renderElement(sup)))
                : this.createMathMLElement("mrow", null, []);
            return this.createMathMLElement("msubsup", null, [baseElem, subElem, supElem]);
        }
        renderMmlGroup(elem) {
            const tagName = resolveGroupTag(elem.props?.position, elem.props?.verticalJustification);
            return this.renderContainerNS(elem, ns.mathML, tagName);
        }
        renderMmlPreSubSuper(elem) {
            const children = [];
            const grouped = keyBy(elem.children, x => x.type);
            const sup = grouped[DomType.MmlSuperArgument];
            const sub = grouped[DomType.MmlSubArgument];
            const supElem = sup ? this.createMathMLElement("mo", null, asArray(this.renderElement(sup))) : null;
            const subElem = sub ? this.createMathMLElement("mo", null, asArray(this.renderElement(sub))) : null;
            const stubElem = this.createMathMLElement("mo", null);
            children.push(this.createMathMLElement("msubsup", null, [stubElem, subElem, supElem]));
            children.push(...this.renderElements(grouped[DomType.MmlBase].children));
            return this.createMathMLElement("mrow", null, children);
        }
        renderMmlGroupChar(elem) {
            const tagName = elem.props.verticalJustification === "bot" ? "mover" : "munder";
            const result = this.renderContainerNS(elem, ns.mathML, tagName);
            if (elem.props.char) {
                result.appendChild(this.createMathMLElement("mo", null, [elem.props.char]));
            }
            return result;
        }
        renderMmlBar(elem) {
            const style = {};
            switch (elem.props.position) {
                case "top":
                    style.textDecoration = "overline";
                    break;
                case "bottom":
                    style.textDecoration = "underline";
                    break;
            }
            return this.renderContainerNS(elem, ns.mathML, "mrow", { style });
        }
        renderMmlRun(elem) {
            return this.toHTML(elem, ns.mathML, "ms");
        }
        renderMllList(elem) {
            const children = this.renderElements(elem.children).map(x => this.createMathMLElement("mtr", null, [
                this.createMathMLElement("mtd", null, [x])
            ]));
            return this.toHTML(elem, ns.mathML, "mtable", children);
        }
        toH(elem, ns, tagName, children = null) {
            const { "$lang": rawLang, ...style } = elem.cssStyle ?? {};
            const lang = isValidBcp47LanguageTag(rawLang) ? rawLang : undefined;
            this.applyThemeColorSideband(style);
            const className = cx(elem.className, elem.styleName && this.processStyleName(elem.styleName));
            return { ns, tagName, className, lang, style, children: children ?? this.renderElements(elem.children) };
        }
        applyThemeColorSideband(style) {
            const palette = this.document?.themePart?.theme?.colorScheme?.colors ?? null;
            const keys = Object.keys(style);
            for (const k of keys) {
                if (!k.startsWith('$themeColor-'))
                    continue;
                const targetProp = k.slice('$themeColor-'.length);
                const placeholder = style[k];
                delete style[k];
                const ref = parseThemeColorReference(placeholder);
                if (!ref)
                    continue;
                const resolved = sanitizeCssColor(resolveColour(ref, palette));
                if (!resolved)
                    continue;
                const existing = style[targetProp];
                if (existing == null) {
                    style[targetProp] = resolved;
                    continue;
                }
                if (targetProp === 'border' || targetProp.startsWith('border-')) {
                    const parts = existing.split(/\s+/);
                    if (parts.length >= 3) {
                        style[targetProp] = `${parts[0]} ${parts[1]} ${resolved}`;
                    }
                    else {
                        style[targetProp] = resolved;
                    }
                    continue;
                }
                style[targetProp] = resolved;
            }
        }
        toHTML(elem, ns, tagName, children = null) {
            return this.h(this.toH(elem, ns, tagName, children));
        }
        findStyle(styleName) {
            return styleName && this.styleMap?.[styleName];
        }
        numberingClass(id, lvl) {
            return this.renderSessionId
                ? `${this.className}-num-${id}-${lvl}-${this.renderSessionId}`
                : `${this.className}-num-${id}-${lvl}`;
        }
        tabStopClass() {
            return `${this.className}-tab-stop`;
        }
        styleToString(selectors, values, cssText = null) {
            let result = `${selectors} {\r\n`;
            for (const key in values) {
                if (key.startsWith('$'))
                    continue;
                result += `  ${key}: ${values[key]};\r\n`;
            }
            if (cssText)
                result += cssText;
            return result + "}\r\n";
        }
        numberingCounter(id, lvl) {
            return this.renderSessionId
                ? `${this.className}-num-${id}-${lvl}-${this.renderSessionId}`
                : `${this.className}-num-${id}-${lvl}`;
        }
        levelTextToContent(text, suff, id, numformat, isLgl = false) {
            const suffMap = {
                "tab": "\\9",
                "space": "\\a0",
            };
            const parts = [];
            let last = 0;
            const re = /%\d+/g;
            let m;
            while ((m = re.exec(text)) !== null) {
                if (m.index > last) {
                    parts.push(`"${escapeCssStringContent(text.slice(last, m.index))}"`);
                }
                const lvl = parseInt(m[0].substring(1), 10) - 1;
                const fmt = isLgl ? "decimal" : numformat;
                parts.push(`counter(${this.numberingCounter(id, lvl)}, ${fmt})`);
                last = re.lastIndex;
            }
            if (last < text.length) {
                parts.push(`"${escapeCssStringContent(text.slice(last))}"`);
            }
            const suffToken = suffMap[suff];
            if (suffToken) {
                parts.push(`"${suffToken}"`);
            }
            return parts.length > 0 ? parts.join(' ') : '""';
        }
        numFormatToCssValue(format) {
            var mapping = {
                none: "none",
                bullet: "disc",
                decimal: "decimal",
                lowerLetter: "lower-alpha",
                upperLetter: "upper-alpha",
                lowerRoman: "lower-roman",
                upperRoman: "upper-roman",
                decimalZero: "decimal-leading-zero",
                aiueo: "katakana",
                aiueoFullWidth: "katakana",
                chineseCounting: "simp-chinese-informal",
                chineseCountingThousand: "simp-chinese-informal",
                chineseLegalSimplified: "simp-chinese-formal",
                chosung: "hangul-consonant",
                ideographDigital: "cjk-ideographic",
                ideographTraditional: "cjk-heavenly-stem",
                ideographLegalTraditional: "trad-chinese-formal",
                ideographZodiac: "cjk-earthly-branch",
                iroha: "katakana-iroha",
                irohaFullWidth: "katakana-iroha",
                japaneseCounting: "japanese-informal",
                japaneseDigitalTenThousand: "cjk-decimal",
                japaneseLegal: "japanese-formal",
                thaiNumbers: "thai",
                koreanCounting: "korean-hangul-formal",
                koreanDigital: "korean-hangul-formal",
                koreanDigital2: "korean-hanja-informal",
                hebrew1: "hebrew",
                hebrew2: "hebrew",
                hindiNumbers: "devanagari",
                ganada: "hangul",
                taiwaneseCounting: "cjk-ideographic",
                taiwaneseCountingThousand: "cjk-ideographic",
                taiwaneseDigital: "cjk-decimal",
            };
            return mapping[format] ?? 'decimal';
        }
        refreshTabStops() {
            if (!this.options.experimental)
                return;
            setTimeout(() => {
                const pixelToPoint = computePixelToPoint();
                for (let tab of this.currentTabs) {
                    updateTabStop(tab.span, tab.stops, this.defaultTabSize, pixelToPoint);
                }
            }, 500);
        }
        createElementNS(ns, tagName, props, children) {
            return this.h({ ns, tagName, children, ...props });
        }
        createElement(tagName, props, children) {
            return this.createElementNS(ns.html, tagName, props, children);
        }
        createMathMLElement(tagName, props, children) {
            return this.createElementNS(ns.mathML, tagName, props, children);
        }
        createSvgElement(tagName, props, children) {
            return this.createElementNS(ns.svg, tagName, props, children);
        }
        later(func) {
            this.postRenderTasks.push(func);
        }
        finalizeChangesRendering(result) {
            const c = this.className;
            const opts = this.options.changes ?? {};
            if (opts.changeBar !== false) {
                for (const el of this.changeElements) {
                    const block = this.findBlockAncestor(el);
                    if (!block)
                        continue;
                    block.classList.add(`${c}-change-bar`);
                    if (!block.style.color) {
                        const match = Array.from(el.classList).find(n => n.startsWith(`${c}-change-author-`));
                        if (match)
                            block.classList.add(match);
                    }
                }
            }
            if (opts.legend !== false && this.changeAuthorIndex.size > 0) {
                const legend = this.buildLegend();
                if (legend) {
                    const wrapper = this.findWrapper(result);
                    if (wrapper) {
                        wrapper.insertBefore(legend, wrapper.firstChild);
                    }
                    else if (result.length) {
                        const insertAt = result.findIndex(n => n.nodeName !== "STYLE" && n.nodeType === 1);
                        if (insertAt >= 0)
                            result.splice(insertAt, 0, legend);
                        else
                            result.push(legend);
                    }
                }
            }
            this.extendSidebarWithChanges();
        }
        extendSidebarWithChanges() {
            const c = this.className;
            const opts = this.options.changes ?? {};
            if (opts.sidebarCards === false)
                return;
            if (!this.useSidebar || !this.sidebarContainer)
                return;
            const content = this.sidebarContainer.querySelector(`.${c}-sidebar-content`);
            if (!content)
                return;
            const seen = new Set();
            const unique = this.changeMeta.filter(m => {
                if (!m.id || seen.has(m.id))
                    return false;
                seen.add(m.id);
                return true;
            });
            for (const meta of unique) {
                const card = this.buildRevisionCard(meta);
                content.appendChild(card);
                if (meta.id)
                    this.revisionCardElements.set(meta.id, card);
            }
        }
        buildRevisionCard(meta) {
            const c = this.className;
            const opts = this.options.changes ?? {};
            const authorIdxClass = meta.author && opts.colorByAuthor !== false
                ? `${c}-change-author-${this.getAuthorIndex(meta.author)}`
                : "";
            const headerChildren = [
                this.h({ tagName: "span", className: `${c}-comment-author ${authorIdxClass}`, children: [meta.author ?? "Unknown"] }),
                this.h({ tagName: "span", className: `${c}-comment-date`, children: [meta.date ? new Date(meta.date).toLocaleString() : ""] }),
                this.h({ tagName: "span", className: `${c}-revision-kind`, children: [this.kindLabel(meta.kind)] }),
            ];
            const card = this.h({
                tagName: "div",
                className: `${c}-sidebar-comment ${c}-revision-card`,
                children: [
                    this.h({ tagName: "div", className: `${c}-comment-header`, children: headerChildren }),
                    this.h({ tagName: "div", className: `${c}-comment-body`, children: [meta.summary] }),
                ]
            });
            card.addEventListener("click", () => {
                meta.el.scrollIntoView({ behavior: "smooth", block: "center" });
            });
            return card;
        }
        kindLabel(kind) {
            switch (kind) {
                case "insertion": return "Inserted";
                case "deletion": return "Deleted";
                case "move": return "Moved";
                case "formatting": return "Formatted";
                case "paragraphMark": return "Paragraph mark";
                case "rowInsertion": return "Row added";
                case "rowDeletion": return "Row removed";
            }
        }
        findBlockAncestor(el) {
            let cur = el.parentElement;
            while (cur) {
                const tag = cur.tagName;
                if (tag === "P" || tag === "LI" || tag === "TR" || tag === "H1" || tag === "H2" ||
                    tag === "H3" || tag === "H4" || tag === "H5" || tag === "H6") {
                    return cur;
                }
                if (tag === "SECTION" || tag === "BODY" || tag === "ARTICLE")
                    return null;
                cur = cur.parentElement;
            }
            return null;
        }
        findWrapper(result) {
            const wrapperClass = `${this.className}-wrapper`;
            for (const node of result) {
                if (node instanceof HTMLElement && node.classList.contains(wrapperClass)) {
                    return node;
                }
            }
            return null;
        }
        buildLegend() {
            const c = this.className;
            const items = [
                this.h({ tagName: "span", className: `${c}-legend-label`, children: ["Changes by:"] })
            ];
            const authors = [...this.changeAuthorIndex.entries()].sort((a, b) => a[1] - b[1]);
            for (const [author, idx] of authors) {
                items.push(this.h({
                    tagName: "span",
                    className: `${c}-legend-item`,
                    children: [
                        this.h({ tagName: "span", className: `${c}-legend-swatch ${c}-change-author-${idx}`, style: { background: "currentColor" } }),
                        author
                    ]
                }));
            }
            return this.h({ tagName: "div", className: `${c}-legend`, children: items });
        }
    }
    HtmlRenderer.CHANGE_PALETTE_SIZE = 8;
    HtmlRenderer.OLE_PROGID_LABELS = {
        'Equation.3': 'Equation',
        'Equation.2': 'Equation',
        'Excel.Sheet.12': 'Excel spreadsheet',
        'Excel.Sheet.8': 'Excel spreadsheet',
        'Excel.SheetBinaryMacroEnabled.12': 'Excel spreadsheet',
        'Excel.SheetMacroEnabled.12': 'Excel spreadsheet',
        'Excel.Chart.8': 'Excel chart',
        'PowerPoint.Show.12': 'PowerPoint slide',
        'PowerPoint.Show.8': 'PowerPoint slide',
        'Word.Document.12': 'Word document',
        'Word.Document.8': 'Word document',
        'Package': 'Embedded file',
        'AcroExch.Document': 'PDF document',
        'AcroExch.Document.DC': 'PDF document',
        'Visio.Drawing.15': 'Visio drawing',
        'Visio.Drawing.11': 'Visio drawing',
    };
    function findParent(elem, type) {
        var parent = elem.parent;
        while (parent != null && parent.type != type)
            parent = parent.parent;
        return parent;
    }

    const defaultMeasure = (el) => {
        const win = el.ownerDocument.defaultView;
        const cs = win?.getComputedStyle(el);
        const rect = el.getBoundingClientRect();
        const width = (cs ? parseFloat(cs.width) : 0) || rect.width || 0;
        const height = (cs ? parseFloat(cs.height) : 0) || rect.height || 0;
        const minHeight = cs ? parseFloat(cs.minHeight) || 0 : 0;
        return { width, height, minHeight };
    };
    const VISUAL_PAGE_MARKER$1 = 'data-docxjs-visual-page';
    function applyVisualPageBreaks(bodyContainer, options = {}, measureFn = defaultMeasure) {
        const className = options.className ?? 'docx';
        const slack = options.slack ?? 1.1;
        const sections = Array.from(bodyContainer.querySelectorAll(`section.${className}`));
        let inserted = 0;
        for (const section of sections) {
            if (section.hasAttribute(VISUAL_PAGE_MARKER$1))
                continue;
            const subPages = splitSection(section, measureFn, slack);
            if (subPages.length > 1) {
                inserted += subPages.length - 1;
                redistributeFootnotes(subPages);
            }
        }
        return inserted;
    }
    function splitSection(section, measureFn, slack) {
        const { height, minHeight } = measureFn(section);
        const pageHeight = minHeight > 0 ? minHeight : 0;
        if (pageHeight <= 0)
            return [section];
        if (height <= pageHeight * slack)
            return [section];
        const article = section.querySelector(':scope > article');
        if (!article)
            return [section];
        const headers = Array.from(section.querySelectorAll(':scope > header'));
        const footers = Array.from(section.querySelectorAll(':scope > footer'));
        const children = Array.from(article.children);
        if (children.length === 0)
            return [section];
        const articleTopOffset = offsetWithinSection(article, section);
        const headerHeight = headers.reduce((sum, h) => sum + measureFn(h).height, 0);
        const footerHeight = footers.reduce((sum, f) => sum + measureFn(f).height, 0);
        const subPages = [section];
        let currentArticle = article;
        let currentTop = articleTopOffset;
        let runningHeight = 0;
        let currentSection = section;
        const roomForCurrent = () => pageHeight - currentTop - footerHeight;
        for (let i = 0; i < children.length; i++) {
            const child = children[i];
            const { height: ch } = measureFn(child);
            let willOverflow = ch > (roomForCurrent() - runningHeight);
            if (willOverflow && child.tagName === 'TABLE') {
                const room = roomForCurrent() - runningHeight;
                const tail = splitTableAtRowBoundary(child, room, measureFn);
                if (tail) {
                    children.splice(i + 1, 0, tail);
                    const { height: newCh } = measureFn(child);
                    runningHeight += newCh;
                    continue;
                }
                willOverflow = true;
            }
            if (willOverflow && runningHeight > 0) {
                const newSection = cloneSectionShell(currentSection);
                for (const h of headers)
                    newSection.appendChild(cloneChromeForRepeat(h));
                const newArticle = cloneArticleShell(currentArticle);
                newSection.appendChild(newArticle);
                for (const f of footers)
                    newSection.appendChild(cloneChromeForRepeat(f));
                currentSection.parentNode.insertBefore(newSection, currentSection.nextSibling);
                subPages.push(newSection);
                for (let j = i; j < children.length; j++) {
                    newArticle.appendChild(children[j]);
                }
                currentSection = newSection;
                currentArticle = newArticle;
                currentTop = headerHeight;
                runningHeight = ch;
                continue;
            }
            runningHeight += ch;
        }
        return subPages;
    }
    function redistributeFootnotes(subPages) {
        const original = subPages[0];
        const originalOls = Array.from(original.querySelectorAll(':scope > ol'));
        if (originalOls.length === 0)
            return;
        const refsBySubPage = subPages.map((page) => {
            const ids = new Set();
            const sups = page.querySelectorAll('article [data-footnote-id]');
            for (const sup of Array.from(sups)) {
                const id = sup.dataset.footnoteId;
                if (id)
                    ids.add(id);
            }
            return ids;
        });
        for (const originalOl of originalOls) {
            const firstLi = originalOl.querySelector(':scope > li');
            if (firstLi && firstLi.id && /^docx-endnote/i.test(firstLi.id)) {
                continue;
            }
            const lis = Array.from(originalOl.children);
            const targetOls = new Map();
            targetOls.set(original, originalOl);
            for (const li of lis) {
                const id = li.dataset.footnoteId;
                if (!id)
                    continue;
                let ownerIdx = -1;
                for (let p = 0; p < subPages.length; p++) {
                    if (refsBySubPage[p].has(id)) {
                        ownerIdx = p;
                        break;
                    }
                }
                if (ownerIdx <= 0)
                    continue;
                const owner = subPages[ownerIdx];
                let ol = targetOls.get(owner);
                if (!ol) {
                    ol = originalOl.cloneNode(false);
                    ol.removeAttribute('id');
                    owner.appendChild(ol);
                    targetOls.set(owner, ol);
                }
                ol.appendChild(li);
            }
            if (originalOl.children.length === 0) {
                originalOl.remove();
            }
            let cumulative = 0;
            for (const page of subPages) {
                const ol = targetOls.get(page);
                if (!ol)
                    continue;
                if (cumulative > 0) {
                    ol.setAttribute('start', String(cumulative + 1));
                }
                cumulative += ol.children.length;
            }
        }
    }
    function offsetWithinSection(child, ancestor, measureFn) {
        const cRect = child.getBoundingClientRect();
        const aRect = ancestor.getBoundingClientRect();
        const delta = cRect.top - aRect.top;
        if (Number.isFinite(delta) && delta >= 0)
            return delta;
        return 0;
    }
    function cloneSectionShell(source) {
        const shell = source.cloneNode(false);
        shell.removeAttribute('id');
        shell.setAttribute(VISUAL_PAGE_MARKER$1, '');
        return shell;
    }
    function cloneArticleShell(source) {
        const shell = source.cloneNode(false);
        shell.removeAttribute('id');
        return shell;
    }
    function splitTableAtRowBoundary(table, room, measureFn) {
        if (room <= 0)
            return null;
        const thead = table.querySelector(':scope > thead');
        const tbodies = Array.from(table.querySelectorAll(':scope > tbody'));
        const rows = tbodies.flatMap(tb => Array.from(tb.children).filter(c => c.tagName === 'TR'));
        if (rows.length === 0)
            return null;
        const theadHeight = thead ? measureFn(thead).height : 0;
        let consumed = theadHeight;
        let cutIndex = -1;
        for (let i = 0; i < rows.length; i++) {
            const rh = measureFn(rows[i]).height;
            if (consumed + rh > room)
                break;
            consumed += rh;
            cutIndex = i;
        }
        if (cutIndex < 0)
            return null;
        if (cutIndex === rows.length - 1)
            return null;
        const nextRow = rows[cutIndex + 1];
        if (nextRow && nextRow.hasAttribute('data-cant-split')) {
            let safe = cutIndex;
            while (safe >= 0) {
                const candidate = rows[safe + 1];
                if (!candidate || !candidate.hasAttribute('data-cant-split'))
                    break;
                safe--;
            }
            if (safe >= 0) {
                cutIndex = safe;
            }
        }
        const tail = table.cloneNode(false);
        tail.removeAttribute('id');
        const colgroup = table.querySelector(':scope > colgroup');
        if (colgroup) {
            const colClone = colgroup.cloneNode(true);
            colClone.removeAttribute('id');
            tail.appendChild(colClone);
        }
        if (thead) {
            const theadClone = thead.cloneNode(true);
            theadClone.removeAttribute('id');
            for (const el of Array.from(theadClone.querySelectorAll('[id]'))) {
                el.removeAttribute('id');
            }
            tail.appendChild(theadClone);
        }
        const tailBody = (tbodies[0] ?? table).cloneNode(false);
        tailBody.removeAttribute('id');
        for (let i = cutIndex + 1; i < rows.length; i++) {
            tailBody.appendChild(rows[i]);
        }
        tail.appendChild(tailBody);
        for (const tb of tbodies) {
            if (!tb.querySelector(':scope > tr'))
                tb.remove();
        }
        return tail;
    }
    function cloneChromeForRepeat(source) {
        const clone = source.cloneNode(true);
        clone.removeAttribute('id');
        for (const el of Array.from(clone.querySelectorAll('[id]'))) {
            el.removeAttribute('id');
        }
        return clone;
    }

    const STYLE_MARKER = 'data-docxjs-thumbnails';
    const VISUAL_PAGE_MARKER = 'data-docxjs-visual-page';
    function findScrollingAncestor(el) {
        let cur = el?.parentElement ?? null;
        while (cur) {
            const cs = cur.ownerDocument.defaultView?.getComputedStyle(cur);
            if (cs) {
                const oy = cs.overflowY;
                if ((oy === 'auto' || oy === 'scroll') && cur.scrollHeight > cur.clientHeight) {
                    return cur;
                }
            }
            cur = cur.parentElement;
        }
        return null;
    }
    function ensureStyle(doc, className, activeClassName) {
        const head = doc.head;
        if (!head)
            return;
        if (head.querySelector(`style[${STYLE_MARKER}]`))
            return;
        const style = doc.createElement('style');
        style.setAttribute(STYLE_MARKER, '');
        style.textContent = `
.${className}-thumbnail {
    display: flex;
    flex-direction: column;
    align-items: center;
    margin: 0.5rem auto;
    cursor: pointer;
    outline: none;
}
.${className}-thumbnail:focus-visible .${className}-thumbnail-preview {
    box-shadow: 0 0 0 2px #4a90e2, 0 0 10px rgba(0, 0, 0, 0.5);
}
.${className}-thumbnail-preview {
    overflow: hidden;
    background: white;
    box-shadow: 0 0 10px rgba(0, 0, 0, 0.5);
    box-sizing: content-box;
    border: 2px solid transparent;
    position: relative;
}
.${className}-thumbnail-label {
    font-size: 0.75rem;
    color: white;
    margin-top: 0.25rem;
    text-align: center;
    line-height: 1.2;
}
.${activeClassName} .${className}-thumbnail-preview {
    border-color: #4a90e2;
}
`;
        head.appendChild(style);
    }
    function measure(el, win) {
        const cs = win?.getComputedStyle(el);
        const rect = el.getBoundingClientRect();
        const width = (cs ? parseFloat(cs.width) : 0) || rect.width || 0;
        const height = (cs ? parseFloat(cs.height) : 0) || rect.height || 0;
        const minHeight = cs ? parseFloat(cs.minHeight) || 0 : 0;
        return { width, height, minHeight };
    }
    function singlePage(section, win) {
        const { width, height, minHeight } = measure(section, win);
        const pageHeight = minHeight > 0 ? minHeight : height;
        return [{
                section, scrollTarget: section,
                topOffset: 0, pageWidth: width, pageHeight,
            }];
    }
    function paginateSection(section, win) {
        const { width, height, minHeight } = measure(section, win);
        const pageHeight = minHeight > 0 ? minHeight : height;
        const pageWidth = width;
        if (pageHeight <= 0 || height <= 0) {
            return [{
                    section, scrollTarget: section,
                    topOffset: 0, pageWidth, pageHeight,
                }];
        }
        const pageCount = Math.max(1, Math.ceil(height / pageHeight));
        if (pageCount === 1) {
            return [{
                    section, scrollTarget: section,
                    topOffset: 0, pageWidth, pageHeight,
                }];
        }
        const cs = win?.getComputedStyle(section);
        if (cs && cs.position === 'static') {
            section.style.position = 'relative';
        }
        const pages = [];
        for (let i = 0; i < pageCount; i++) {
            let anchor = section.querySelector(`[data-docxjs-page-anchor="${i}"]`);
            if (!anchor) {
                anchor = section.ownerDocument.createElement('div');
                anchor.setAttribute('data-docxjs-page-anchor', String(i));
                anchor.setAttribute('aria-hidden', 'true');
                anchor.style.cssText = [
                    'position:absolute',
                    `top:${i * pageHeight}px`,
                    'left:0',
                    `width:${pageWidth}px`,
                    `height:${pageHeight}px`,
                    'pointer-events:none',
                    'visibility:hidden',
                ].join(';');
                section.appendChild(anchor);
            }
            pages.push({
                section, scrollTarget: anchor,
                topOffset: i * pageHeight,
                pageWidth, pageHeight,
            });
        }
        return pages;
    }
    function renderThumbnails(mainContainer, thumbnailContainer, options) {
        const width = options?.width ?? 120;
        const showPageNumbers = options?.showPageNumbers ?? true;
        const className = options?.className ?? 'docx';
        const activeClassName = options?.activeClassName ?? `${className}-thumbnail-active`;
        const doc = thumbnailContainer.ownerDocument;
        const win = doc.defaultView;
        ensureStyle(mainContainer.ownerDocument, className, activeClassName);
        thumbnailContainer.innerHTML = '';
        const sections = Array.from(mainContainer.querySelectorAll(`section.${className}`));
        const splitterRan = mainContainer.querySelector(`section[${VISUAL_PAGE_MARKER}]`) !== null;
        const pages = [];
        for (const section of sections) {
            const sectionPages = splitterRan
                ? singlePage(section, win)
                : paginateSection(section, win);
            for (const p of sectionPages) {
                pages.push(p);
            }
        }
        const pairs = [];
        for (let i = 0; i < pages.length; i++) {
            const { section, scrollTarget, topOffset, pageWidth, pageHeight } = pages[i];
            const pageNum = i + 1;
            const thumb = doc.createElement('div');
            thumb.className = `${className}-thumbnail`;
            thumb.setAttribute('role', 'button');
            thumb.setAttribute('tabindex', '0');
            thumb.setAttribute('aria-label', `Go to page ${pageNum}`);
            thumb.dataset.page = String(pageNum);
            const preview = doc.createElement('div');
            preview.className = `${className}-thumbnail-preview`;
            const clone = section.cloneNode(true);
            clone.setAttribute('aria-hidden', 'true');
            clone.removeAttribute('id');
            clone.style.boxShadow = 'none';
            clone.style.margin = '0';
            clone.style.flexShrink = '0';
            let scale = 1;
            let previewHeight = 0;
            if (pageWidth > 0) {
                scale = width / pageWidth;
                previewHeight = pageHeight * scale;
            }
            preview.style.width = `${width}px`;
            if (previewHeight > 0) {
                preview.style.height = `${previewHeight}px`;
            }
            const translate = -topOffset * scale;
            clone.style.transform = `translateY(${translate}px) scale(${scale})`;
            clone.style.transformOrigin = '0 0';
            preview.appendChild(clone);
            thumb.appendChild(preview);
            if (showPageNumbers) {
                const label = doc.createElement('div');
                label.className = `${className}-thumbnail-label`;
                label.textContent = String(pageNum);
                thumb.appendChild(label);
            }
            const goTo = () => {
                scrollTarget.scrollIntoView({ behavior: 'smooth', block: 'start' });
            };
            thumb.addEventListener('click', goTo);
            thumb.addEventListener('keydown', (ev) => {
                const ke = ev;
                if (ke.key === 'Enter' || ke.key === ' ') {
                    ke.preventDefault();
                    goTo();
                }
            });
            thumbnailContainer.appendChild(thumb);
            pairs.push({ scrollTarget, thumb });
        }
        let observer = null;
        const IO = win?.IntersectionObserver ??
            globalThis.IntersectionObserver;
        if (IO && pairs.length > 0) {
            const scrollRoot = findScrollingAncestor(mainContainer);
            const visibility = new Map();
            observer = new IO((entries) => {
                for (const entry of entries) {
                    visibility.set(entry.target, entry.intersectionRatio);
                }
                let bestIdx = -1;
                let bestRatio = -1;
                for (let i = 0; i < pairs.length; i++) {
                    const r = visibility.get(pairs[i].scrollTarget) ?? 0;
                    if (r > bestRatio) {
                        bestRatio = r;
                        bestIdx = i;
                    }
                }
                for (let i = 0; i < pairs.length; i++) {
                    pairs[i].thumb.classList.toggle(activeClassName, i === bestIdx && bestRatio > 0);
                }
            }, {
                root: scrollRoot,
                rootMargin: '-45% 0px -45% 0px',
                threshold: [0, 0.01, 0.5, 1],
            });
            for (const { scrollTarget } of pairs)
                observer.observe(scrollTarget);
        }
        return {
            dispose() {
                if (observer) {
                    observer.disconnect();
                    observer = null;
                }
                thumbnailContainer.innerHTML = '';
                for (const section of sections) {
                    section.querySelectorAll('[data-docxjs-page-anchor]').forEach(n => n.remove());
                }
            },
        };
    }

    const defaultOptions = {
        ignoreHeight: false,
        ignoreWidth: false,
        ignoreFonts: false,
        breakPages: true,
        debug: false,
        experimental: false,
        className: "docx",
        inWrapper: true,
        hideWrapperOnPrint: false,
        trimXmlDeclaration: true,
        ignoreLastRenderedPageBreak: true,
        renderHeaders: true,
        renderFooters: true,
        renderFootnotes: true,
        renderEndnotes: true,
        useBase64URL: false,
        renderChanges: false,
        renderComments: false,
        showProtectionBadge: false,
        experimentalPageBreaks: false,
        responsive: false,
        emitDocumentProps: false,
        comments: {
            sidebar: true,
            highlight: true,
            layout: 'anchored',
        },
        changes: {
            show: false,
            showInsertions: true,
            showDeletions: true,
            showMoves: true,
            showFormatting: true,
            colorByAuthor: true,
            changeBar: true,
            legend: true,
            sidebarCards: true,
        },
        h: h
    };
    function mergeOptions(userOptions) {
        const ops = { ...defaultOptions, ...userOptions };
        if (userOptions?.renderChanges && userOptions?.changes?.show === undefined) {
            ops.changes = { ...defaultOptions.changes, ...userOptions.changes, show: true };
        }
        return ops;
    }
    function parseAsync(data, userOptions) {
        const ops = mergeOptions(userOptions);
        return WordDocument.load(data, new DocumentParser(ops), ops);
    }
    async function renderDocument(document, userOptions) {
        const ops = mergeOptions(userOptions);
        const renderer = new HtmlRenderer();
        return await renderer.render(document, ops);
    }
    async function renderAsync(data, bodyContainer, styleContainer, userOptions) {
        const doc = await parseAsync(data, userOptions);
        const nodes = await renderDocument(doc, userOptions);
        styleContainer ?? (styleContainer = bodyContainer);
        styleContainer.innerHTML = "";
        bodyContainer.innerHTML = "";
        for (let n of nodes) {
            const c = n.nodeName === "STYLE" ? styleContainer : bodyContainer;
            c.appendChild(n);
        }
        const ops = mergeOptions(userOptions);
        if (ops.experimentalPageBreaks) {
            applyVisualPageBreaks(bodyContainer, { className: ops.className });
        }
        return doc;
    }

    exports.applyVisualPageBreaks = applyVisualPageBreaks;
    exports.classNameOfCnfStyle = classNameOfCnfStyle;
    exports.defaultOptions = defaultOptions;
    exports.escapeCssStringContent = escapeCssStringContent;
    exports.isSafeCssIdent = isSafeCssIdent;
    exports.isSafeCustomXmlXPath = isSafeCustomXmlXPath;
    exports.isSafeHyperlinkHref = isSafeHyperlinkHref;
    exports.keyBy = keyBy;
    exports.layoutTreemap = layoutTreemap;
    exports.mergeDeep = mergeDeep;
    exports.parseAsync = parseAsync;
    exports.parseFieldInstruction = parseFieldInstruction;
    exports.renderAsync = renderAsync;
    exports.renderDocument = renderDocument;
    exports.renderThumbnails = renderThumbnails;
    exports.safeEvaluateXPath = safeEvaluateXPath;
    exports.sanitizeCssColor = sanitizeCssColor;
    exports.sanitizeFontFamily = sanitizeFontFamily;
    exports.sanitizeVmlColor = sanitizeVmlColor;

}));
//# sourceMappingURL=docx-preview.js.map
