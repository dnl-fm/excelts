
import BaseXform from '../base-xform.ts';

class BooleanXform extends BaseXform {
  constructor(options) {
    super();

    this.tag = options.tag;
    this.attr = options.attr;
  }

  render(xmlStream, model) {
    if (model) {
      xmlStream.openNode(this.tag);
      xmlStream.closeNode();
    }
  }

  parseOpen(node) {
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