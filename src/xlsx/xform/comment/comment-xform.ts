/**
  <comment ref="B1" authorId="0">
    <text>
      <r>
        <rPr>
          <b/>
          <sz val="9"/>
          <rFont val="宋体"/>
          <charset val="134"/>
        </rPr>
        <t>51422:</t>
      </r>
      <r>
        <rPr>
          <sz val="9"/>
          <rFont val="宋体"/>
          <charset val="134"/>
        </rPr>
        <t xml:space="preserve">&#10;test</t>
      </r>
    </text>
  </comment>
 */

import RichTextXform from '../strings/rich-text-xform.ts';
import BaseXform from '../base-xform.ts';
import type { SaxNode, XmlStreamWriter } from '../xform-types.ts';

interface CommentModel {
  type?: string;
  ref?: string;
  note?: {
    texts: unknown[];
  };
  [key: string]: unknown;
}

class CommentXform extends BaseXform {
  declare model: CommentModel;
  _richTextXform: RichTextXform | null;

  constructor(model?: CommentModel) {
    super();
    if (model) {
      this.model = model;
    }
    this._richTextXform = null;
    this.parser = undefined;
  }

  get tag(): string {
    return 'r';
  }

  get richTextXform(): RichTextXform {
    if (!this._richTextXform) {
      this._richTextXform = new RichTextXform();
    }
    return this._richTextXform;
  }

  render(xmlStream: XmlStreamWriter, model?: CommentModel): void {
    const m = model || this.model;

    xmlStream.openNode('comment', {
      ref: m.ref,
      authorId: 0,
    });
    xmlStream.openNode('text');
    if (m && m.note && m.note.texts) {
      m.note.texts.forEach(text => {
        this.richTextXform.render(xmlStream, text);
      });
    }
    xmlStream.closeNode();
    xmlStream.closeNode();
  }

  parseOpen(node: SaxNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'comment':
        this.model = {
          type: 'note',
          note: {
            texts: [],
          },
          ...(node.attributes as Record<string, string>),
        };
        return true;
      case 'r':
        this.parser = this.richTextXform;
        this.parser.parseOpen(node);
        return true;
      default:
        return false;
    }
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    switch (name) {
      case 'comment':
        return false;
      case 'r':
        this.model.note!.texts.push(this.parser!.model);
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

export default CommentXform;
