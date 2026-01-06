/* eslint-disable max-classes-per-file */

import BaseXform from '../../base-xform.ts';
import CompositeXform from '../../composite-xform.ts';

type ExtRefModel = {
  x14Id?: string;
};

type XmlStreamLike = {
  openNode: (tag: string, attrs?: Record<string, unknown>) => void;
  closeNode: () => void;
  leafNode: (tag: string, attrs: unknown, content: unknown) => void;
};

class X14IdXform extends BaseXform {
  declare model: string;

  get tag() {
    return 'x14:id';
  }

  render(xmlStream: XmlStreamLike, model: string) {
    xmlStream.leafNode(this.tag, null, model);
  }

  parseOpen() {
    this.model = '';
  }

  parseText(text: string) {
    this.model += text;
  }

  parseClose(name: string) {
    return name !== this.tag;
  }
}

class ExtXform extends CompositeXform {
  declare model: ExtRefModel;
  declare idXform: X14IdXform;

  constructor() {
    super();

    this.map = {
      'x14:id': (this.idXform = new X14IdXform()),
    };
  }

  get tag() {
    return 'ext';
  }

  render(xmlStream: XmlStreamLike, model: ExtRefModel) {
    xmlStream.openNode(this.tag, {
      uri: '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}',
      'xmlns:x14': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
    });

    this.idXform.render(xmlStream, model.x14Id!);

    xmlStream.closeNode();
  }

  createNewModel(): ExtRefModel {
    return {};
  }

  onParserClose(_name: string, parser: { model: string }) {
    this.model.x14Id = parser.model;
  }
}

class ExtLstRefXform extends CompositeXform {
  declare model: ExtRefModel;
  declare extXform: ExtXform;

  constructor() {
    super();
    this.extXform = new ExtXform();
    this.map = {
      ext: this.extXform,
    };
  }

  get tag() {
    return 'extLst';
  }

  render(xmlStream: XmlStreamLike, model: ExtRefModel) {
    xmlStream.openNode(this.tag);
    this.extXform.render(xmlStream, model);
    xmlStream.closeNode();
  }

  createNewModel(): ExtRefModel {
    return {};
  }

  onParserClose(_name: string, parser: { model: ExtRefModel }) {
    Object.assign(this.model, parser.model);
  }
}

export default ExtLstRefXform;