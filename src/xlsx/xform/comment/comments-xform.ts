import XmlStream from '../../../utils/xml-stream.ts';
import BaseXform from '../base-xform.ts';
import CommentXform from './comment-xform.ts';
import type {XmlStreamWriter} from '../xform-types.ts';

type CommentModel = {
  ref?: string;
  author?: string;
  text?: unknown;
  [key: string]: unknown;
};

type CommentsModel = {
  comments: CommentModel[];
  [key: string]: unknown;
};

class CommentsXform extends BaseXform {
  static COMMENTS_ATTRIBUTES = {
    xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  };

  declare model: CommentsModel | undefined;
  declare map: {
    comment: CommentXform;
  };

  constructor() {
    super();
    this.map = {
      comment: new CommentXform(),
    };
    this.parser = null;
  }

  render(xmlStream: XmlStreamWriter, model?: CommentsModel) {
    model = model || this.model;
    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode('comments', CommentsXform.COMMENTS_ATTRIBUTES);

    // authors
    // TODO: support authors properly
    xmlStream.openNode('authors');
    xmlStream.leafNode('author', null, 'Author');
    xmlStream.closeNode();

    // comments
    xmlStream.openNode('commentList');
    model.comments.forEach(comment => {
      this.map.comment.render(xmlStream, comment);
    });
    xmlStream.closeNode();
    xmlStream.closeNode();
  }

  parseOpen(node: {name: string; attributes: Record<string, string>}) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'commentList':
        this.model = {
          comments: [],
        };
        return true;
      case 'comment':
        this.parser = this.map.comment;
        this.parser.parseOpen(node);
        return true;
      default:
        return false;
    }
  }

  parseText(text: string) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string) {
    switch (name) {
      case 'commentList':
        return false;
      case 'comment':
        this.model!.comments.push(this.parser!.model as CommentModel);
        this.parser = undefined;
        return true;
      default:
        if (this.parser) {
          this.parser.parseClose(name);
        }
        return true;
    }
  }
}

export default CommentsXform;
