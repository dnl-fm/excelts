import BaseXform from '../../base-xform.ts';

// Use Web Crypto API for UUID generation (browser-compatible)
function randomUUID(): string {
  return crypto.randomUUID();
}
import CompositeXform from '../../composite-xform.ts';
import DatabarExtXform from './databar-ext-xform.ts';
import IconSetExtXform from './icon-set-ext-xform.ts';
import type {XmlStreamWriter, SaxNode} from '../../xform-types.ts';

const extIcons: Record<string, boolean> = {
  '3Triangles': true,
  '3Stars': true,
  '5Boxes': true,
};

type CfRuleExtModel = {
  type?: string;
  x14Id?: string;
  priority?: number;
  custom?: boolean;
  iconSet?: string;
  gradient?: boolean;
};

class CfRuleExtXform extends CompositeXform {
  databarXform: DatabarExtXform;
  iconSetXform: IconSetExtXform;

  constructor() {
    super();

    this.map = {
      'x14:dataBar': (this.databarXform = new DatabarExtXform()),
      'x14:iconSet': (this.iconSetXform = new IconSetExtXform()),
    };
  }

  get tag() {
    return 'x14:cfRule';
  }

  static isExt(rule: CfRuleExtModel) {
    // is this rule primitive?
    if (rule.type === 'dataBar') {
      return DatabarExtXform.isExt(rule);
    }
    if (rule.type === 'iconSet') {
      if (rule.custom || (rule.iconSet && extIcons[rule.iconSet])) {
        return true;
      }
    }
    return false;
  }

  prepare(model: CfRuleExtModel, _options?: unknown) {
    if (CfRuleExtXform.isExt(model)) {
      model.x14Id = `{${randomUUID()}}`.toUpperCase();
    }
  }

  render(xmlStream: XmlStreamWriter, model: CfRuleExtModel) {
    if (!CfRuleExtXform.isExt(model)) {
      return;
    }

    switch (model.type) {
      case 'dataBar':
        this.renderDataBar(xmlStream, model);
        break;
      case 'iconSet':
        this.renderIconSet(xmlStream, model);
        break;
    }
  }

  renderDataBar(xmlStream: XmlStreamWriter, model: CfRuleExtModel) {
    xmlStream.openNode(this.tag, {
      type: 'dataBar',
      id: model.x14Id,
    });

    this.databarXform.render(xmlStream, model as any);

    xmlStream.closeNode();
  }

  renderIconSet(xmlStream: XmlStreamWriter, model: CfRuleExtModel) {
    xmlStream.openNode(this.tag, {
      type: 'iconSet',
      priority: model.priority,
      id: model.x14Id || `{${randomUUID()}}`,
    });

    this.iconSetXform.render(xmlStream, model as any);

    xmlStream.closeNode();
  }

  createNewModel({attributes}: SaxNode) {
    const model: {type?: string; x14Id?: string; priority?: number} = {
      type: attributes.type,
      x14Id: attributes.id,
    };
    // Only include priority if it was present in the XML
    if (attributes.priority !== undefined) {
      model.priority = BaseXform.toIntValue(attributes.priority);
    }
    return model;
  }

  onParserClose(_name: string, parser: BaseXform) {
    Object.assign(this.model as object, parser.model as object);
  }
}

export default CfRuleExtXform;
