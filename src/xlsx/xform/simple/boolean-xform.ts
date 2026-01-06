import BaseXform from '../base-xform.ts';
import type {SaxNode} from '../xform-types.ts';

type BooleanXformOptions = {
  tag: string;
  attr?: string;
};

class BooleanXform extends BaseXform {
  declare tag: string;
  attr?: string;
  // Override model to allow boolean
  declare model: boolean | undefined;

  constructor(options: BooleanXformOptions) {
    super();

    this.tag = options.tag;
    this.attr = options.attr;
  }

  render(xmlStream: {openNode: (tag: string) => void; closeNode: () => void}, model: boolean) {
    if (model) {
      xmlStream.openNode(this.tag);
      xmlStream.closeNode();
    }
  }

  parseOpen(node: SaxNode) {
    if (node.name === this.tag) {
      // Check the val attribute: val="0" or val="false" means false
      // Otherwise (no val attr, val="1", val="true") means true
      const val = this.attr && node.attributes && node.attributes[this.attr];
      this.model = val !== '0' && val !== 'false';
    }
  }

  parseText() {}

  parseClose() {
    return false;
  }
}

export default BooleanXform;
