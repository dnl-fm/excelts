
//   <t xml:space="preserve"> is </t>

import BaseXform from '../base-xform.ts';

class TextXform extends BaseXform {
  declare model: string | undefined;
  private _text: string[] = [];

  get tag() {
    return 't';
  }

  render(xmlStream: { openNode: (tag: string) => void; addAttribute: (name: string, value: string) => void; writeText: (text: string) => void; closeNode: () => void }, model: string) {
    xmlStream.openNode('t');
    if (/^\s|\n|\s$/.test(model)) {
      xmlStream.addAttribute('xml:space', 'preserve');
    }
    xmlStream.writeText(model);
    xmlStream.closeNode();
  }

  private computeModel(): string {
    return this._text
      .join('')
      .replace(/_x([0-9A-F]{4})_/g, (_$0, $1) => String.fromCharCode(parseInt($1, 16)));
  }

  parseOpen(node: { name: string }) {
    switch (node.name) {
      case 't':
        this._text = [];
        return true;
      default:
        return false;
    }
  }

  parseText(text: string) {
    this._text.push(text);
  }

  parseClose() {
    // Compute and store model when element closes
    this.model = this.computeModel();
    return false;
  }
}

export default TextXform;