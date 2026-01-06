
import BaseXform from '../base-xform.ts';
import type { SaxNode, XmlStreamWriter } from '../xform-types.ts';

interface FloatXformOptions {
  tag: string;
  attr?: string;
  attrs?: Record<string, string>;
}

class FloatXform extends BaseXform {
  declare tag: string;
  attr?: string;
  attrs?: Record<string, string>;
  text?: string[];

  constructor(options: FloatXformOptions) {
    super();

    this.tag = options.tag;
    this.attr = options.attr;
    this.attrs = options.attrs;
  }

  render(xmlStream: XmlStreamWriter, model: number): void {
    if (model !== undefined) {
      xmlStream.openNode(this.tag);
      if (this.attrs) {
        xmlStream.addAttributes(this.attrs);
      }
      if (this.attr) {
        xmlStream.addAttribute(this.attr, model);
      } else {
        xmlStream.writeText(String(model));
      }
      xmlStream.closeNode();
    }
  }

  parseOpen(node: SaxNode): boolean {
    if (node.name === this.tag) {
      if (this.attr) {
        (this as unknown as { model: number }).model = parseFloat(node.attributes[this.attr] as string);
      } else {
        this.text = [];
      }
    }
    return true;
  }

  parseText(text: string): void {
    if (!this.attr && this.text) {
      this.text.push(text);
    }
  }

  parseClose(): boolean {
    if (!this.attr && this.text) {
      (this as unknown as { model: number }).model = parseFloat(this.text.join(''));
    }
    return false;
  }
}

export default FloatXform;
