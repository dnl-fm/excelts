import BaseCellAnchorXform from './base-cell-anchor-xform.ts';
import StaticXform from '../static-xform.ts';
import CellPositionXform from './cell-position-xform.ts';
import PicXform from './pic-xform.ts';
import type {XmlStreamWriter} from '../xform-types.ts';

type TwoCellAnchorModel = {
  range: {
    editAs?: string;
    tl?: unknown;
    br?: unknown;
  };
  picture?: unknown;
  medium?: unknown;
  [key: string]: unknown;
};

class TwoCellAnchorXform extends BaseCellAnchorXform {
  declare model: TwoCellAnchorModel | undefined;
  declare map: {
    'xdr:from': CellPositionXform;
    'xdr:to': CellPositionXform;
    'xdr:pic': PicXform;
    'xdr:clientData': StaticXform;
  };

  constructor() {
    super();

    this.map = {
      'xdr:from': new CellPositionXform({tag: 'xdr:from'}),
      'xdr:to': new CellPositionXform({tag: 'xdr:to'}),
      'xdr:pic': new PicXform(),
      'xdr:clientData': new StaticXform({tag: 'xdr:clientData'}),
    };
  }

  get tag() {
    return 'xdr:twoCellAnchor';
  }

  prepare(model: TwoCellAnchorModel, options: unknown) {
    this.map['xdr:pic'].prepare(model.picture, options);
  }

  render(xmlStream: XmlStreamWriter, model: TwoCellAnchorModel) {
    xmlStream.openNode(this.tag, {editAs: model.range.editAs || 'oneCell'});

    this.map['xdr:from'].render(xmlStream, model.range.tl);
    this.map['xdr:to'].render(xmlStream, model.range.br);
    this.map['xdr:pic'].render(xmlStream, model.picture);
    this.map['xdr:clientData'].render(xmlStream, {});

    xmlStream.closeNode();
  }

  parseClose(name: string) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        this.model!.range.tl = this.map['xdr:from'].model;
        this.model!.range.br = this.map['xdr:to'].model;
        this.model!.picture = this.map['xdr:pic'].model;
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  }

  reconcile(model: TwoCellAnchorModel, options: unknown) {
    model.medium = this.reconcilePicture(model.picture, options);
  }
}

export default TwoCellAnchorXform;
