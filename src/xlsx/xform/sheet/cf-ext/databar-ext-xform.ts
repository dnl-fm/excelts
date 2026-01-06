import BaseXform from '../../base-xform.ts';
import CompositeXform from '../../composite-xform.ts';
import ColorXform from '../../style/color-xform.ts';
import CfvoExtXform from './cfvo-ext-xform.ts';
import type {XmlStreamWriter, SaxNode} from '../../xform-types.ts';

type CfvoModel = {
  type?: string;
  value?: string | number;
};

type ColorModel = {
  argb?: string;
  theme?: number;
  tint?: number;
};

type DatabarExtModel = {
  cfvo: CfvoModel[];
  minLength?: number;
  maxLength?: number;
  border?: boolean;
  gradient?: boolean;
  negativeBarColorSameAsPositive?: boolean;
  negativeBarBorderColorSameAsPositive?: boolean;
  axisPosition?: unknown;
  direction?: unknown;
  borderColor?: ColorModel;
  negativeBorderColor?: ColorModel;
  negativeFillColor?: ColorModel;
  axisColor?: ColorModel;
  [key: string]: unknown;
};

class DatabarExtXform extends CompositeXform {
  cfvoXform: CfvoExtXform;
  borderColorXform: ColorXform;
  negativeBorderColorXform: ColorXform;
  negativeFillColorXform: ColorXform;
  axisColorXform: ColorXform;

  declare model: DatabarExtModel | undefined;

  constructor() {
    super();

    this.map = {
      'x14:cfvo': (this.cfvoXform = new CfvoExtXform()),
      'x14:borderColor': (this.borderColorXform = new ColorXform('x14:borderColor')),
      'x14:negativeBorderColor': (this.negativeBorderColorXform = new ColorXform(
        'x14:negativeBorderColor'
      )),
      'x14:negativeFillColor': (this.negativeFillColorXform = new ColorXform(
        'x14:negativeFillColor'
      )),
      'x14:axisColor': (this.axisColorXform = new ColorXform('x14:axisColor')),
    };
  }

  static isExt(rule: {gradient?: boolean}) {
    // not all databars need ext
    // TODO: refine this
    return !rule.gradient;
  }

  get tag() {
    return 'x14:dataBar';
  }

  render(xmlStream: XmlStreamWriter, model: DatabarExtModel) {
    xmlStream.openNode(this.tag, {
      minLength: BaseXform.toIntAttribute(model.minLength, 0, true),
      maxLength: BaseXform.toIntAttribute(model.maxLength, 100, true),
      border: BaseXform.toBoolAttribute(model.border, false),
      gradient: BaseXform.toBoolAttribute(model.gradient, true),
      negativeBarColorSameAsPositive: BaseXform.toBoolAttribute(
        model.negativeBarColorSameAsPositive,
        true
      ),
      negativeBarBorderColorSameAsPositive: BaseXform.toBoolAttribute(
        model.negativeBarBorderColorSameAsPositive,
        true
      ),
      axisPosition: BaseXform.toAttribute(model.axisPosition, 'auto'),
      direction: BaseXform.toAttribute(model.direction, 'leftToRight'),
    });

    model.cfvo.forEach(cfvo => {
      if (cfvo.type) {
        this.cfvoXform.render(xmlStream, cfvo as { type: string; value?: string | number });
      }
    });

    this.borderColorXform.render(xmlStream, model.borderColor);
    this.negativeBorderColorXform.render(xmlStream, model.negativeBorderColor);
    this.negativeFillColorXform.render(xmlStream, model.negativeFillColor);
    this.axisColorXform.render(xmlStream, model.axisColor);

    xmlStream.closeNode();
  }

  createNewModel({attributes}: SaxNode): DatabarExtModel {
    return {
      cfvo: [],
      minLength: BaseXform.toIntValue(attributes.minLength, 0),
      maxLength: BaseXform.toIntValue(attributes.maxLength, 100),
      border: BaseXform.toBoolValue(attributes.border, false),
      gradient: BaseXform.toBoolValue(attributes.gradient, true),
      negativeBarColorSameAsPositive: BaseXform.toBoolValue(
        attributes.negativeBarColorSameAsPositive,
        true
      ),
      negativeBarBorderColorSameAsPositive: BaseXform.toBoolValue(
        attributes.negativeBarBorderColorSameAsPositive,
        true
      ),
      axisPosition: BaseXform.toStringValue(attributes.axisPosition, 'auto'),
      direction: BaseXform.toStringValue(attributes.direction, 'leftToRight'),
    };
  }

  onParserClose(name: string, parser: BaseXform) {
    const [, prop] = name.split(':');
    switch (prop) {
      case 'cfvo':
        this.model!.cfvo.push(parser.model as CfvoModel);
        break;

      default:
        this.model![prop] = parser.model;
        break;
    }
  }
}

export default DatabarExtXform;
