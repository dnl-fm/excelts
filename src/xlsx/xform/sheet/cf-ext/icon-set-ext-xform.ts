import BaseXform from '../../base-xform.ts';
import CompositeXform from '../../composite-xform.ts';
import CfvoExtXform from './cfvo-ext-xform.ts';
import CfIconExtXform from './cf-icon-ext-xform.ts';
import type {XmlStreamWriter, SaxNode} from '../../xform-types.ts';

type CfvoModel = {
  type?: string;
  value?: string | number;
};

type CfIconModel = {
  iconSet?: string;
  iconId?: number;
};

type IconSetExtModel = {
  cfvo: CfvoModel[];
  icons?: CfIconModel[];
  iconSet?: unknown;
  reverse?: boolean;
  showValue?: boolean;
  [key: string]: unknown;
};

class IconSetExtXform extends CompositeXform {
  cfvoXform: CfvoExtXform;
  cfIconXform: CfIconExtXform;

  declare model: IconSetExtModel | undefined;

  constructor() {
    super();

    this.map = {
      'x14:cfvo': (this.cfvoXform = new CfvoExtXform()),
      'x14:cfIcon': (this.cfIconXform = new CfIconExtXform()),
    };
  }

  get tag() {
    return 'x14:iconSet';
  }

  render(xmlStream: XmlStreamWriter, model: IconSetExtModel) {
    xmlStream.openNode(this.tag, {
      iconSet: BaseXform.toStringAttribute(model.iconSet),
      reverse: BaseXform.toBoolAttribute(model.reverse, false),
      showValue: BaseXform.toBoolAttribute(model.showValue, true),
      custom: BaseXform.toBoolAttribute(model.icons, false),
    });

    model.cfvo.forEach(cfvo => {
      if (cfvo.type) {
        this.cfvoXform.render(xmlStream, cfvo as { type: string; value?: string | number });
      }
    });

    if (model.icons) {
      model.icons.forEach((icon, i) => {
        icon.iconId = i;
        this.cfIconXform.render(xmlStream, icon);
      });
    }

    xmlStream.closeNode();
  }

  createNewModel({attributes}: SaxNode): IconSetExtModel {
    return {
      cfvo: [],
      iconSet: BaseXform.toStringValue(attributes.iconSet, '3TrafficLights'),
      reverse: BaseXform.toBoolValue(attributes.reverse, false),
      showValue: BaseXform.toBoolValue(attributes.showValue, true),
    };
  }

  onParserClose(name: string, parser: BaseXform) {
    const [, prop] = name.split(':');
    switch (prop) {
      case 'cfvo':
        this.model!.cfvo.push(parser.model as CfvoModel);
        break;

      case 'cfIcon':
        if (!this.model!.icons) {
          this.model!.icons = [];
        }
        this.model!.icons.push(parser.model as CfIconModel);
        break;

      default:
        this.model![prop] = parser.model;
        break;
    }
  }
}

export default IconSetExtXform;
