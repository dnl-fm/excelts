import colCache from '../../../utils/col-cache.ts';
import XmlStream from '../../../utils/xml-stream.ts';
import BaseXform from '../base-xform.ts';
import TwoCellAnchorXform from './two-cell-anchor-xform.ts';
import OneCellAnchorXform from './one-cell-anchor-xform.ts';
import type {XmlStreamWriter, SaxNode} from '../xform-types.ts';

type DrawingModel = {
  anchors?: Array<Record<string, unknown>>;
};

function getAnchorType(model: Record<string, unknown>) {
  const range = typeof model.range === 'string' 
    ? colCache.decode(model.range)
    : model.range as { br?: unknown } | undefined;

  return range && 'br' in range ? 'xdr:twoCellAnchor' : 'xdr:oneCellAnchor';
}

class DrawingXform extends BaseXform {
  static DRAWING_ATTRIBUTES: Record<string, string> = {
    'xmlns:xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
  };

  declare model: DrawingModel | undefined;
  declare map: {
    'xdr:twoCellAnchor': TwoCellAnchorXform;
    'xdr:oneCellAnchor': OneCellAnchorXform;
  };

  constructor() {
    super();

    this.map = {
      'xdr:twoCellAnchor': new TwoCellAnchorXform(),
      'xdr:oneCellAnchor': new OneCellAnchorXform(),
    };
  }

  prepare(model: DrawingModel, _options?: unknown) {
    model.anchors.forEach((item, index) => {
      item.anchorType = getAnchorType(item);
      const anchor = this.map[item.anchorType as keyof typeof this.map];
      anchor.prepare(item as any, {index});
    });
  }

  get tag() {
    return 'xdr:wsDr';
  }

  render(xmlStream: XmlStreamWriter, model: DrawingModel) {
    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode(this.tag, DrawingXform.DRAWING_ATTRIBUTES);

    model.anchors.forEach(item => {
      const anchor = this.map[item.anchorType as keyof typeof this.map];
      anchor.render(xmlStream, item as any);
    });

    xmlStream.closeNode();
  }

  parseOpen(node: SaxNode) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          anchors: [],
        };
        break;
      default:
        this.parser = this.map[node.name as keyof typeof this.map];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  }

  parseText(text: string) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model!.anchors.push(this.parser.model as Record<string, unknown>);
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  }

  reconcile(model: DrawingModel, options: unknown) {
    model.anchors.forEach(anchor => {
      const range = anchor.range as { br?: unknown } | undefined;
      if (range?.br) {
        this.map['xdr:twoCellAnchor'].reconcile(anchor as any, options);
      } else {
        this.map['xdr:oneCellAnchor'].reconcile(anchor as any, options);
      }
    });
  }
}

export default DrawingXform;
