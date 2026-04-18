import { Part } from "../common/part";
import { OpenXmlPackage } from "../common/open-xml-package";
import { DocumentParser } from "../document-parser";
import { WmlComment } from "./elements";
import { CommentsExtended } from "./comments-extended-part";
import { keyBy } from "../utils";

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

        const extMap = keyBy(extendedComments, x => x.paraId);
        const paraIdToComment: Record<string, WmlComment> = {};

        for (const comment of this.comments) {
            if (comment.paraId) {
                paraIdToComment[comment.paraId] = comment;
                const ext = extMap[comment.paraId];
                if (ext) {
                    comment.done = ext.done;
                }
            }
        }

        for (const ext of extendedComments) {
            if (ext.paraIdParent) {
                const child = paraIdToComment[ext.paraId];
                const parent = paraIdToComment[ext.paraIdParent];
                if (child && parent) {
                    child.parentCommentId = parent.id;
                    parent.replies.push(child);
                }
            }
        }

        this.topLevelComments = this.comments.filter(c => !c.parentCommentId);
    }
}