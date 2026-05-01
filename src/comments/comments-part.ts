import { Part } from "../common/part";
import { OpenXmlPackage } from "../common/open-xml-package";
import { DocumentParser } from "../document-parser";
import { WmlComment } from "./elements";
import { CommentsExtended } from "./comments-extended-part";
import { keyBy } from "../utils";

// paraId comes from an untrusted DOCX; reject values like "__proto__" that would
// otherwise poison the prototype chain of the lookup map.
const SAFE_PARA_ID = /^[A-Za-z0-9_-]+$/;

export class CommentsPart extends Part {
    protected _documentParser: DocumentParser;

    comments: WmlComment[]
    commentMap: Record<string, WmlComment>;
    topLevelComments: WmlComment[] = [];

    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
        super(pkg, path);
        this._documentParser = parser;
    }

	parseXml(root: Element) {
        this.comments = this._documentParser.parseComments(root);
		this.commentMap = keyBy(this.comments, x => x.id);
    }

    buildThreading(extendedComments: CommentsExtended[]) {
        if (!extendedComments || extendedComments.length === 0) {
            this.topLevelComments = [...this.comments];
            return;
        }

        const extMap = new Map<string, CommentsExtended>();
        for (const ext of extendedComments) {
            if (ext.paraId && SAFE_PARA_ID.test(ext.paraId)) {
                extMap.set(ext.paraId, ext);
            }
        }

        const paraIdToComment = new Map<string, WmlComment>();
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
