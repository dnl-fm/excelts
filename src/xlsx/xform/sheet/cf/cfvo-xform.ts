
import BaseXform from '../../base-xform.ts';

class CfvoXform extends BaseXform {
  get tag() {
    return 'cfvo';
  }

  render(xmlStream, model) {
    xmlStream.leafNode(this.tag, {
      type: model.type,
      val: model.value,
    });
  }

  parseOpen(node: { name: string; attributes: Record<string, string> }): boolean {
    const attrs = node.attributes;
    const model: { type: string; value?: number } = { type: attrs.type };
    if (attrs.val !== undefined) {
      model.value = BaseXform.toFloatValue(attrs.val);
    }
    this.model = model;
    return true;
  }

  parseClose(name) {
    return name !== this.tag;
  }
}

export default CfvoXform;