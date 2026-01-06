
import BaseXform from '../base-xform.ts';

type VmlTextboxModel = {
  inset?: number[];
};

type VmlTextboxRenderModel = {
  note?: {
    margins?: {
      inset?: number[] | string;
    };
  };
};

class VmlTextboxXform extends BaseXform {
  declare model: VmlTextboxModel | undefined;

  get tag() {
    return 'v:textbox';
  }

  conversionUnit(value: number | string, multiple: number, unit: string) {
    return `${parseFloat(String(value)) * Number(multiple.toFixed(2))}${unit}`;
  }

  reverseConversionUnit(inset: string | undefined) {
    return (inset || '').split(',').map(margin => {
      return Number(parseFloat(this.conversionUnit(parseFloat(margin), 0.1, '')).toFixed(2));
    });
  }

  render(xmlStream: {openNode: (tag: string, attrs: Record<string, unknown>) => void; leafNode: (tag: string, attrs: Record<string, unknown>) => void; closeNode: () => void}, model: VmlTextboxRenderModel) {
    const attributes: {style: string; inset?: string} = {
      style: 'mso-direction-alt:auto',
    };
    if (model && model.note) {
      const marginInset = model.note.margins?.inset;
      let insetStr: string | undefined;
      if (Array.isArray(marginInset)) {
        insetStr = marginInset
          .map(margin => {
            return this.conversionUnit(margin, 10, 'mm');
          })
          .join(',');
      } else if (typeof marginInset === 'string') {
        insetStr = marginInset;
      }
      if (insetStr) {
        attributes.inset = insetStr;
      }
    }
    xmlStream.openNode('v:textbox', attributes);
    xmlStream.leafNode('div', {style: 'text-align:left'});
    xmlStream.closeNode();
  }

  parseOpen(node: {name: string; attributes: {inset?: string}}) {
    switch (node.name) {
      case this.tag:
        this.model = {
          inset: this.reverseConversionUnit(node.attributes.inset),
        };
        return true;
      default:
        return true;
    }
  }

  parseText() {}

  parseClose(name: string) {
    switch (name) {
      case this.tag:
        return false;
      default:
        return true;
    }
  }
}

export default VmlTextboxXform;