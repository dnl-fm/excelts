
import BaseXform from '../base-xform.ts';
import type { SaxNode, XmlStreamWriter } from '../xform-types.ts';

class UnderlineXform extends BaseXform {
  static Attributes: Record<string, Record<string, string>>;

  constructor(model?: boolean | string) {
    super();

    if (model !== undefined) {
      (this as unknown as { model: boolean | string }).model = model;
    }
  }

  get tag(): string {
    return 'u';
  }

  render(xmlStream: XmlStreamWriter, model?: boolean | string): void {
    const m = model ?? this.model;

    if (m === true) {
      xmlStream.leafNode('u', null);
    } else if (typeof m === 'string') {
      const attr = UnderlineXform.Attributes[m];
      if (attr) {
        xmlStream.leafNode('u', attr);
      }
    }
  }

  parseOpen(node: SaxNode): boolean {
    if (node.name === 'u') {
      (this as unknown as { model: boolean | string }).model = node.attributes.val || true;
    }
    return true;
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }
}

UnderlineXform.Attributes = {
  single: {},
  double: {val: 'double'},
  singleAccounting: {val: 'singleAccounting'},
  doubleAccounting: {val: 'doubleAccounting'},
};

export default UnderlineXform;