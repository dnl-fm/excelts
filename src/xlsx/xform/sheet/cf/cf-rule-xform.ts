

import BaseXform from '../../base-xform.ts';
import CompositeXform from '../../composite-xform.ts';
import Range from '../../../../doc/range.ts';
import DatabarXform from './databar-xform.ts';
import ExtLstRefXform from './ext-lst-ref-xform.ts';
import FormulaXform from './formula-xform.ts';
import ColorScaleXform from './color-scale-xform.ts';
import IconSetXform from './icon-set-xform.ts';
import type { SaxNode, XmlStreamWriter } from '../../xform-types.ts';

interface CfRuleModel {
  type?: string;
  operator?: string;
  dxfId?: number;
  priority?: number;
  timePeriod?: string;
  percent?: boolean;
  bottom?: boolean;
  rank?: number;
  aboveAverage?: boolean;
  formulae?: string[];
  text?: string;
  ref?: string;
  iconSet?: string;
  custom?: boolean;
  // Properties for databar, colorScale, iconSet child xforms
  cfvo?: unknown[];
  color?: unknown;
  x14Id?: string;
  reverse?: boolean;
  showValue?: boolean;
  [key: string]: unknown;
}

const extIcons: Record<string, boolean> = {
  '3Triangles': true,
  '3Stars': true,
  '5Boxes': true,
};

const getTextFormula = (model: CfRuleModel): string | undefined => {
  if (model.formulae && model.formulae[0]) {
    return model.formulae[0];
  }

  const range = new Range(model.ref);
  const {tl} = range;
  switch (model.operator) {
    case 'containsText':
      return `NOT(ISERROR(SEARCH("${model.text}",${tl})))`;
    case 'containsBlanks':
      return `LEN(TRIM(${tl}))=0`;
    case 'notContainsBlanks':
      return `LEN(TRIM(${tl}))>0`;
    case 'containsErrors':
      return `ISERROR(${tl})`;
    case 'notContainsErrors':
      return `NOT(ISERROR(${tl}))`;
    default:
      return undefined;
  }
};

const getTimePeriodFormula = (model: CfRuleModel): string | undefined => {
  if (model.formulae && model.formulae[0]) {
    return model.formulae[0];
  }

  const range = new Range(model.ref);
  const {tl} = range;
  switch (model.timePeriod) {
    case 'thisWeek':
      return `AND(TODAY()-ROUNDDOWN(${tl},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${tl},0)-TODAY()<=7-WEEKDAY(TODAY()))`;
    case 'lastWeek':
      return `AND(TODAY()-ROUNDDOWN(${tl},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${tl},0)<(WEEKDAY(TODAY())+7))`;
    case 'nextWeek':
      return `AND(ROUNDDOWN(${tl},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${tl},0)-TODAY()<(15-WEEKDAY(TODAY())))`;
    case 'yesterday':
      return `FLOOR(${tl},1)=TODAY()-1`;
    case 'today':
      return `FLOOR(${tl},1)=TODAY()`;
    case 'tomorrow':
      return `FLOOR(${tl},1)=TODAY()+1`;
    case 'last7Days':
      return `AND(TODAY()-FLOOR(${tl},1)<=6,FLOOR(${tl},1)<=TODAY())`;
    case 'lastMonth':
      return `AND(MONTH(${tl})=MONTH(EDATE(TODAY(),0-1)),YEAR(${tl})=YEAR(EDATE(TODAY(),0-1)))`;
    case 'thisMonth':
      return `AND(MONTH(${tl})=MONTH(TODAY()),YEAR(${tl})=YEAR(TODAY()))`;
    case 'nextMonth':
      return `AND(MONTH(${tl})=MONTH(EDATE(TODAY(),0+1)),YEAR(${tl})=YEAR(EDATE(TODAY(),0+1)))`;
    default:
      return undefined;
  }
};

const opType = (attributes: Record<string, string>): { type: string; operator?: string } => {
  const {type, operator} = attributes;
  switch (type) {
    case 'containsText':
    case 'containsBlanks':
    case 'notContainsBlanks':
    case 'containsErrors':
    case 'notContainsErrors':
      return {
        type: 'containsText',
        operator: type,
      };

    default:
      return {type, operator};
  }
};

class CfRuleXform extends CompositeXform {
  declare model: CfRuleModel;
  databarXform: DatabarXform;
  extLstRefXform: ExtLstRefXform;
  formulaXform: FormulaXform;
  colorScaleXform: ColorScaleXform;
  iconSetXform: IconSetXform;

  constructor() {
    super();

    this.databarXform = new DatabarXform();
    this.extLstRefXform = new ExtLstRefXform();
    this.formulaXform = new FormulaXform();
    this.colorScaleXform = new ColorScaleXform();
    this.iconSetXform = new IconSetXform();

    this.map = {
      dataBar: this.databarXform,
      extLst: this.extLstRefXform,
      formula: this.formulaXform,
      colorScale: this.colorScaleXform,
      iconSet: this.iconSetXform,
    };
  }

  get tag(): string {
    return 'cfRule';
  }

  static isPrimitive(rule: CfRuleModel): boolean {
    // is this rule primitive?
    if (rule.type === 'iconSet') {
      if (rule.custom || extIcons[rule.iconSet]) {
        return false;
      }
    }
    return true;
  }

  render(xmlStream: XmlStreamWriter, model: CfRuleModel): void {
    switch (model.type) {
      case 'expression':
        this.renderExpression(xmlStream, model);
        break;
      case 'cellIs':
        this.renderCellIs(xmlStream, model);
        break;
      case 'top10':
        this.renderTop10(xmlStream, model);
        break;
      case 'aboveAverage':
        this.renderAboveAverage(xmlStream, model);
        break;
      case 'dataBar':
        this.renderDataBar(xmlStream, model);
        break;
      case 'colorScale':
        this.renderColorScale(xmlStream, model);
        break;
      case 'iconSet':
        this.renderIconSet(xmlStream, model);
        break;
      case 'containsText':
        this.renderText(xmlStream, model);
        break;
      case 'timePeriod':
        this.renderTimePeriod(xmlStream, model);
        break;
    }
  }

  renderExpression(xmlStream: XmlStreamWriter, model: CfRuleModel): void {
    xmlStream.openNode(this.tag, {
      type: 'expression',
      dxfId: model.dxfId,
      priority: model.priority,
    });

    this.formulaXform.render(xmlStream, model.formulae![0]);

    xmlStream.closeNode();
  }

  renderCellIs(xmlStream: XmlStreamWriter, model: CfRuleModel): void {
    xmlStream.openNode(this.tag, {
      type: 'cellIs',
      dxfId: model.dxfId,
      priority: model.priority,
      operator: model.operator,
    });

    model.formulae!.forEach(formula => {
      this.formulaXform.render(xmlStream, formula);
    });

    xmlStream.closeNode();
  }

  renderTop10(xmlStream: XmlStreamWriter, model: CfRuleModel): void {
    xmlStream.leafNode(this.tag, {
      type: 'top10',
      dxfId: model.dxfId,
      priority: model.priority,
      percent: BaseXform.toBoolAttribute(model.percent, false),
      bottom: BaseXform.toBoolAttribute(model.bottom, false),
      rank: BaseXform.toIntAttribute(model.rank, 10, true),
    });
  }

  renderAboveAverage(xmlStream: XmlStreamWriter, model: CfRuleModel): void {
    xmlStream.leafNode(this.tag, {
      type: 'aboveAverage',
      dxfId: model.dxfId,
      priority: model.priority,
      aboveAverage: BaseXform.toBoolAttribute(model.aboveAverage, true),
    });
  }

  renderDataBar(xmlStream: XmlStreamWriter, model: CfRuleModel): void {
    xmlStream.openNode(this.tag, {
      type: 'dataBar',
      priority: model.priority,
    });

    // Model has cfvo array for databar rules
    if (model.cfvo) {
      this.databarXform.render(xmlStream, model as { cfvo: unknown[]; color?: unknown });
    }
    if (model.x14Id) {
      this.extLstRefXform.render(xmlStream, model as { x14Id: string });
    }

    xmlStream.closeNode();
  }

  renderColorScale(xmlStream: XmlStreamWriter, model: CfRuleModel): void {
    xmlStream.openNode(this.tag, {
      type: 'colorScale',
      priority: model.priority,
    });

    // Model has cfvo and color arrays for colorScale rules
    if (model.cfvo && model.color) {
      this.colorScaleXform.render(xmlStream, model as { cfvo: unknown[]; color: unknown[] });
    }

    xmlStream.closeNode();
  }

  renderIconSet(xmlStream: XmlStreamWriter, model: CfRuleModel): void {
    // iconset is all primitive or all extLst
    if (!CfRuleXform.isPrimitive(model)) {
      return;
    }

    xmlStream.openNode(this.tag, {
      type: 'iconSet',
      priority: model.priority,
    });

    // Model has cfvo array for iconSet rules
    if (model.cfvo) {
      this.iconSetXform.render(xmlStream, model as { cfvo: unknown[] });
    }

    xmlStream.closeNode();
  }

  renderText(xmlStream: XmlStreamWriter, model: CfRuleModel): void {
    xmlStream.openNode(this.tag, {
      type: model.operator,
      dxfId: model.dxfId,
      priority: model.priority,
      operator: BaseXform.toStringAttribute(model.operator, 'containsText'),
    });

    const formula = getTextFormula(model);
    if (formula) {
      this.formulaXform.render(xmlStream, formula);
    }

    xmlStream.closeNode();
  }

  renderTimePeriod(xmlStream: XmlStreamWriter, model: CfRuleModel): void {
    xmlStream.openNode(this.tag, {
      type: 'timePeriod',
      dxfId: model.dxfId,
      priority: model.priority,
      timePeriod: model.timePeriod,
    });

    const formula = getTimePeriodFormula(model);
    if (formula) {
      this.formulaXform.render(xmlStream, formula);
    }

    xmlStream.closeNode();
  }

  createNewModel(node: SaxNode): CfRuleModel {
    const attributes = node.attributes as Record<string, string>;
    const model: CfRuleModel = {
      ...opType(attributes),
      priority: BaseXform.toIntValue(attributes.priority),
    };
    if (attributes.dxfId !== undefined) model.dxfId = BaseXform.toIntValue(attributes.dxfId);
    if (attributes.timePeriod !== undefined) model.timePeriod = attributes.timePeriod;
    if (attributes.percent !== undefined) model.percent = BaseXform.toBoolValue(attributes.percent);
    if (attributes.bottom !== undefined) model.bottom = BaseXform.toBoolValue(attributes.bottom);
    if (attributes.rank !== undefined) model.rank = BaseXform.toIntValue(attributes.rank);
    if (attributes.aboveAverage !== undefined) model.aboveAverage = BaseXform.toBoolValue(attributes.aboveAverage);
    return model;
  }

  onParserClose(name: string, parser: BaseXform): void {
    switch (name) {
      case 'dataBar':
      case 'extLst':
      case 'colorScale':
      case 'iconSet':
        // merge parser model with ours
        Object.assign(this.model, parser.model);
        break;

      case 'formula':
        // except - formula is a string and appends to formulae
        this.model.formulae = this.model.formulae || [];
        this.model.formulae.push(parser.model as unknown as string);
        break;
    }
  }
}

export default CfRuleXform;