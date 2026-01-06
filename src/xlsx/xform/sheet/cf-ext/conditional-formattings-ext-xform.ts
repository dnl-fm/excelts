import BaseXform from '../../base-xform.ts';
import CompositeXform from '../../composite-xform.ts';
import CfRuleExtXform from './cf-rule-ext-xform.ts';
import ConditionalFormattingExtXform from './conditional-formatting-ext-xform.ts';
import type {XmlStreamWriter} from '../../xform-types.ts';

type CfRuleExtModel = {
  type?: string;
  x14Id?: string;
  priority?: number;
  custom?: boolean;
  dataBar?: unknown;
  iconSet?: string;
  gradient?: boolean;
  [key: string]: unknown;
};

type ConditionalFormattingExtModel = {
  ref?: string;
  rules: CfRuleExtModel[];
};

type ConditionalFormattingsModel = ConditionalFormattingExtModel[] & {
  hasExtContent?: boolean;
};

type PrepareOptions = {
  styles?: unknown;
};

class ConditionalFormattingsExtXform extends CompositeXform {
  cfXform: ConditionalFormattingExtXform;

  declare model: ConditionalFormattingsModel | undefined;

  constructor() {
    super();

    this.map = {
      'x14:conditionalFormatting': (this.cfXform = new ConditionalFormattingExtXform()),
    };
  }

  get tag() {
    return 'x14:conditionalFormattings';
  }

  hasContent(model: ConditionalFormattingsModel) {
    if (model.hasExtContent === undefined) {
      model.hasExtContent = model.some(cf => cf.rules.some(CfRuleExtXform.isExt));
    }
    return model.hasExtContent;
  }

  prepare(model: ConditionalFormattingsModel, options: PrepareOptions) {
    model.forEach(cf => {
      this.cfXform.prepare(cf, options);
    });
  }

  render(xmlStream: XmlStreamWriter, model: ConditionalFormattingsModel) {
    if (this.hasContent(model)) {
      xmlStream.openNode(this.tag);
      model.forEach(cf => this.cfXform.render(xmlStream, cf));
      xmlStream.closeNode();
    }
  }

  createNewModel(): ConditionalFormattingsModel {
    return [] as ConditionalFormattingsModel;
  }

  onParserClose(_name: string, parser: BaseXform) {
    // model is array of conditional formatting objects
    (this.model as ConditionalFormattingsModel).push(parser.model as ConditionalFormattingExtModel);
  }
}

export default ConditionalFormattingsExtXform;
