

import CompositeXform from '../../composite-xform.ts';
import FExtXform from './f-ext-xform.ts';

export type CfvoExtModel = {
  type: string;
  value?: number | string;
};

type XmlStreamLike = {
  openNode: (tag: string, attrs?: Record<string, unknown>) => void;
  closeNode: () => void;
  leafNode: (tag: string, attrs: unknown, content: unknown) => void;
};

class CfvoExtXform extends CompositeXform {
  declare model: CfvoExtModel;
  declare fExtXform: FExtXform;

  constructor() {
    super();

    this.map = {
      'xm:f': (this.fExtXform = new FExtXform()),
    };
  }

  get tag() {
    return 'x14:cfvo';
  }

  render(xmlStream: XmlStreamLike, model: CfvoExtModel) {
    xmlStream.openNode(this.tag, {
      type: model.type,
    });
    if (model.value !== undefined) {
      this.fExtXform.render(xmlStream, model.value);
    }
    xmlStream.closeNode();
  }

  createNewModel(node: { attributes: Record<string, string> }): CfvoExtModel {
    return {
      type: node.attributes.type,
    };
  }

  onParserClose(name: string, parser: { model: string }) {
    switch (name) {
      case 'xm:f':
        this.model.value = parser.model ? parseFloat(parser.model) : 0;
        break;
    }
  }
}

export default CfvoExtXform;