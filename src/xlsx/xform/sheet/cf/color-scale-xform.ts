

import CompositeXform from '../../composite-xform.ts';
import ColorXform from '../../style/color-xform.ts';
import CfvoXform from './cfvo-xform.ts';
import BaseXform from '../../base-xform.ts';
import type { SaxNode, XmlStreamWriter } from '../../xform-types.ts';

interface ColorScaleModel {
  cfvo: unknown[];
  color: unknown[];
  [key: string]: unknown;
}

class ColorScaleXform extends CompositeXform {
  declare model: ColorScaleModel;
  cfvoXform: CfvoXform;
  colorXform: ColorXform;

  constructor() {
    super();

    this.cfvoXform = new CfvoXform();
    this.colorXform = new ColorXform();
    this.map = {
      cfvo: this.cfvoXform,
      color: this.colorXform,
    };
  }

  get tag(): string {
    return 'colorScale';
  }

  render(xmlStream: XmlStreamWriter, model: ColorScaleModel): void {
    xmlStream.openNode(this.tag);

    model.cfvo.forEach(cfvo => {
      this.cfvoXform.render(xmlStream, cfvo);
    });
    model.color.forEach(color => {
      this.colorXform.render(xmlStream, color);
    });

    xmlStream.closeNode();
  }

  createNewModel(_node: SaxNode): ColorScaleModel {
    return {
      cfvo: [],
      color: [],
    };
  }

  onParserClose(name: string, parser: BaseXform): void {
    (this.model[name] as unknown[]).push(parser.model);
  }
}

export default ColorScaleXform;