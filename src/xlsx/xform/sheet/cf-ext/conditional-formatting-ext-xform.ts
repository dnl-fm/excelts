import BaseXform from '../../base-xform.ts';
import CompositeXform from '../../composite-xform.ts';
import SqRefExtXform from './sqref-ext-xform.ts';
import CfRuleExtXform from './cf-rule-ext-xform.ts';
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

type PrepareOptions = {
  styles?: unknown;
};

class ConditionalFormattingExtXform extends CompositeXform {
  sqRef: SqRefExtXform;
  cfRule: CfRuleExtXform;

  declare model: ConditionalFormattingExtModel | undefined;

  constructor() {
    super();

    this.map = {
      'xm:sqref': (this.sqRef = new SqRefExtXform()),
      'x14:cfRule': (this.cfRule = new CfRuleExtXform()),
    };
  }

  get tag() {
    return 'x14:conditionalFormatting';
  }

  prepare(model: ConditionalFormattingExtModel, options: PrepareOptions) {
    model.rules.forEach(rule => {
      this.cfRule.prepare(rule, options);
    });
  }

  render(xmlStream: XmlStreamWriter, model: ConditionalFormattingExtModel) {
    if (!model.rules.some(CfRuleExtXform.isExt)) {
      return;
    }

    xmlStream.openNode(this.tag, {
      'xmlns:xm': 'http://schemas.microsoft.com/office/excel/2006/main',
    });

    model.rules.filter(CfRuleExtXform.isExt).forEach(rule => this.cfRule.render(xmlStream, rule));

    // for some odd reason, Excel needs the <xm:sqref> node to be after the rules
    this.sqRef.render(xmlStream, model.ref);

    xmlStream.closeNode();
  }

  createNewModel(): ConditionalFormattingExtModel {
    return {
      rules: [],
    };
  }

  onParserClose(name: string, parser: BaseXform) {
    switch (name) {
      case 'xm:sqref':
        this.model!.ref = parser.model as string;
        break;

      case 'x14:cfRule':
        this.model!.rules.push(parser.model as CfRuleExtModel);
        break;
    }
  }
}

export default ConditionalFormattingExtXform;
