
import BaseXform from '../../base-xform.ts';
import type {XmlStreamWriter} from '../../xform-types.ts';

type VmlProtectionConfig = {
  tag: string;
};

class VmlProtectionXform extends BaseXform {
  private _model: VmlProtectionConfig;
  text: string = '';

  constructor(model: VmlProtectionConfig) {
    super();
    this._model = model;
  }

  get tag() {
    return this._model && this._model.tag;
  }

  render(xmlStream: XmlStreamWriter, model: unknown) {
    xmlStream.leafNode(this.tag, null, model);
  }

  parseOpen(node: {name: string}) {
    switch (node.name) {
      case this.tag:
        this.text = '';
        return true;
      default:
        return false;
    }
  }

  parseText(text: string) {
    this.text = text;
  }

  parseClose() {
    return false;
  }
}

export default VmlProtectionXform;