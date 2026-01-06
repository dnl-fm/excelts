
import BaseXform from '../../base-xform.ts';

type XmlStreamLike = {
  leafNode: (tag: string, attrs: unknown, content: unknown) => void;
  openNode?: (tag: string, attrs?: Record<string, unknown>) => void;
  closeNode?: () => void;
};

class FExtXform extends BaseXform {
  declare model: string;

  get tag() {
    return 'xm:f';
  }

  render(xmlStream: XmlStreamLike, model: unknown) {
    xmlStream.leafNode(this.tag, null, model);
  }

  parseOpen() {
    this.model = '';
  }

  parseText(text: string) {
    this.model += text;
  }

  parseClose(name: string) {
    return name !== this.tag;
  }
}

export default FExtXform;