
import BaseXform from '../base-xform.ts';
import type { SaxNode, XmlStreamWriter } from '../xform-types.ts';

interface IntegerXformOptions {
  tag: string;
  attr?: string;
  attrs?: Record<string, string>;
  zero?: boolean;
}

class IntegerXform extends BaseXform {
  declare tag: string;
  attr?: string;
  attrs?: Record<string, string>;
  zero?: boolean;
  text?: string[];

  constructor(options: IntegerXformOptions) {
    super();

    this.tag = options.tag;
    this.attr = options.attr;
    this.attrs = options.attrs;

    // option to render zero
    this.zero = options.zero;
  }

  render(xmlStream: XmlStreamWriter, model: number): void {
    // int is different to float in that zero is not rendered
    if (model || this.zero) {
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
        (this as unknown as { model: number }).model = parseInt(node.attributes[this.attr] as string, 10);
      } else {
        this.text = [];
      }
      return true;
    }
    return false;
  }

  parseText(text: string): void {
    if (!this.attr && this.text) {
      this.text.push(text);
    }
  }

  parseClose(): boolean {
    if (!this.attr && this.text) {
      (this as unknown as { model: number }).model = parseInt(this.text.join('') || '0', 10);
    }
    return false;
  }
}

export default IntegerXform;
