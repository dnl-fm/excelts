import BaseCellAnchorXform from './base-cell-anchor-xform.ts';
import StaticXform from '../static-xform.ts';
import CellPositionXform from './cell-position-xform.ts';
import ExtXform from './ext-xform.ts';
import PicXform from './pic-xform.ts';
import type {XmlStreamWriter} from '../xform-types.ts';

type OneCellAnchorModel = {
  range: {
    editAs?: string;
    tl?: unknown;
    ext?: unknown;
  };
  picture?: unknown;
  medium?: unknown;
  [key: string]: unknown;
};

class OneCellAnchorXform extends BaseCellAnchorXform {
  declare model: OneCellAnchorModel | undefined;
  declare map: {
    'xdr:from': CellPositionXform;
    'xdr:ext': ExtXform;
    'xdr:pic': PicXform;
    'xdr:clientData': StaticXform;
  };

  constructor() {
    super();

    this.map = {
      'xdr:from': new CellPositionXform({tag: 'xdr:from'}),
      'xdr:ext': new ExtXform({tag: 'xdr:ext'}),
      'xdr:pic': new PicXform(),
      'xdr:clientData': new StaticXform({tag: 'xdr:clientData'}),
    };
  }

  get tag() {
    return 'xdr:oneCellAnchor';
  }

  prepare(model: OneCellAnchorModel, options: unknown) {
    this.map['xdr:pic'].prepare(model.picture, options);
  }

  render(xmlStream: XmlStreamWriter, model: OneCellAnchorModel) {
    xmlStream.openNode(this.tag, {editAs: model.range.editAs || 'oneCell'});

    this.map['xdr:from'].render(xmlStream, model.range.tl);
    this.map['xdr:ext'].render(xmlStream, model.range.ext);
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
        this.model!.range.ext = this.map['xdr:ext'].model;
        this.model!.picture = this.map['xdr:pic'].model;
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  }

  reconcile(model: OneCellAnchorModel, options: unknown) {
    model.medium = this.reconcilePicture(model.picture, options);
  }
}

export default OneCellAnchorXform;
