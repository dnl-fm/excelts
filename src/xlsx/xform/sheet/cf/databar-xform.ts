

import CompositeXform from '../../composite-xform.ts';
import ColorXform from '../../style/color-xform.ts';
import CfvoXform from './cfvo-xform.ts';
import BaseXform from '../../base-xform.ts';
import type { XmlStreamWriter } from '../../xform-types.ts';

interface DatabarModel {
  cfvo: unknown[];
  color?: unknown;
  [key: string]: unknown;
}

class DatabarXform extends CompositeXform {
  declare model: DatabarModel;
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
    return 'dataBar';
  }

  render(xmlStream: XmlStreamWriter, model: DatabarModel): void {
    xmlStream.openNode(this.tag);

    model.cfvo.forEach(cfvo => {
      this.cfvoXform.render(xmlStream, cfvo);
    });
    this.colorXform.render(xmlStream, model.color);

    xmlStream.closeNode();
  }

  createNewModel(): DatabarModel {
    return {
      cfvo: [],
    };
  }

  onParserClose(name: string, parser: BaseXform): void {
    switch (name) {
      case 'cfvo':
        this.model.cfvo.push(parser.model);
        break;
      case 'color':
        this.model.color = parser.model;
        break;
    }
  }
}

export default DatabarXform;