
import BaseXform from '../../base-xform.ts';
import type {XmlStreamWriter} from '../../xform-types.ts';

type VmlPositionConfig = {
  tag: string;
};

class VmlPositionXform extends BaseXform {
  private _model: VmlPositionConfig;
  declare model: Record<string, boolean> | undefined;

  constructor(model: VmlPositionConfig) {
    super();
    this._model = model;
  }

  get tag() {
    return this._model && this._model.tag;
  }

  render(xmlStream: XmlStreamWriter, model: string, type: string[]) {
    if (model === type[2]) {
      xmlStream.leafNode(this.tag, null);
    } else if (this.tag === 'x:SizeWithCells' && model === type[1]) {
      xmlStream.leafNode(this.tag, null);
    }
  }

  parseOpen(node: {name: string}) {
    switch (node.name) {
      case this.tag:
        this.model = {};
        this.model[this.tag] = true;
        return true;
      default:
        return false;
    }
  }

  parseText() {}

  parseClose() {
    return false;
  }
}

export default VmlPositionXform;