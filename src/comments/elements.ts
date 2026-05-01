import { DomType, OpenXmlElementBase } from "../document/dom";

export class WmlComment extends OpenXmlElementBase {
	type = DomType.Comment;
    id: string;
	author: string;
	initials: string;
	date: string;
	paraId: string;
	done: boolean = false;
	parentCommentId: string = null;
	replies: WmlComment[] = [];
}

export class WmlCommentReference  extends OpenXmlElementBase {
	type = DomType.CommentReference;
	
	constructor(public id?: string) {
		super();
	}
}

export class WmlCommentRangeStart  extends OpenXmlElementBase {
	type = DomType.CommentRangeStart;
	
	constructor(public id?: string) {
		super();
	}
}
export class WmlCommentRangeEnd  extends OpenXmlElementBase {
	type = DomType.CommentRangeEnd;

	constructor(public id?: string) {
		super();
	}
}