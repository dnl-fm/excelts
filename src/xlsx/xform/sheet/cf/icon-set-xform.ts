

import BaseXform from '../../base-xform.ts';
import CompositeXform from '../../composite-xform.ts';
import CfvoXform from './cfvo-xform.ts';

type CfvoModel = {
  type?: string;
  value?: unknown;
};

type IconSetModel = {
  iconSet?: string;
  reverse?: boolean;
  showValue?: boolean;
  cfvo: CfvoModel[];
};

class IconSetXform extends CompositeXform {
  declare model: IconSetModel;
  declare cfvoXform: CfvoXform;

  constructor() {
    super();

    this.map = {
      cfvo: (this.cfvoXform = new CfvoXform()),
    };
  }

  get tag() {
    return 'iconSet';
  }

  render(xmlStream: { openNode: (tag: string, attrs: Record<string, unknown>) => void; closeNode: () => void }, model: IconSetModel) {
    xmlStream.openNode(this.tag, {
      iconSet: BaseXform.toStringAttribute(model.iconSet, '3TrafficLights'),
      reverse: BaseXform.toBoolAttribute(model.reverse, false),
      showValue: BaseXform.toBoolAttribute(model.showValue, true),
    });

    model.cfvo.forEach(cfvo => {
      this.cfvoXform.render(xmlStream, cfvo);
    });

    xmlStream.closeNode();
  }

  createNewModel({attributes}: { attributes: Record<string, string> }): IconSetModel {
    const model: IconSetModel = {
      iconSet: BaseXform.toStringValue(attributes.iconSet, '3TrafficLights') as string,
      cfvo: [],
    };
    if (attributes.reverse !== undefined) model.reverse = BaseXform.toBoolValue(attributes.reverse);
    if (attributes.showValue !== undefined) model.showValue = BaseXform.toBoolValue(attributes.showValue);
    return model;
  }

  onParserClose(name: string, parser: { model: CfvoModel }) {
    if (name === 'cfvo') {
      this.model.cfvo.push(parser.model);
    }
  }
}

export default IconSetXform;