
import BaseXform from '../base-xform.ts';
import VmlTextboxXform from './vml-textbox-xform.ts';
import VmlClientDataXform from './vml-client-data-xform.ts';
import type { SaxNode, XmlStreamWriter } from '../xform-types.ts';

interface VmlShapeMargins {
  insetmode?: string;
  inset?: string;
}

interface VmlShapeModel {
  margins: VmlShapeMargins;
  anchor: string;
  editAs: string;
  protection: Record<string, unknown>;
  note?: {
    margins?: VmlShapeMargins;
  };
  [key: string]: unknown;
}

class VmlShapeXform extends BaseXform {
  declare model: VmlShapeModel;
  static V_SHAPE_ATTRIBUTES: (model: VmlShapeModel, index: number) => Record<string, unknown>;
  
  constructor() {
    super();
    this.map = {
      'v:textbox': new VmlTextboxXform(),
      'x:ClientData': new VmlClientDataXform(),
    };
  }

  get tag(): string {
    return 'v:shape';
  }

  render(xmlStream: XmlStreamWriter, model: VmlShapeModel, index: number): void {
    xmlStream.openNode('v:shape', VmlShapeXform.V_SHAPE_ATTRIBUTES(model, index));

    xmlStream.leafNode('v:fill', {color2: 'infoBackground [80]'});
    xmlStream.leafNode('v:shadow', {color: 'none [81]', obscured: 't'});
    xmlStream.leafNode('v:path', {'o:connecttype': 'none'});
    this.map['v:textbox'].render(xmlStream, model);
    this.map['x:ClientData'].render(xmlStream, model);

    xmlStream.closeNode();
  }

  parseOpen(node: SaxNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          margins: {
            insetmode: node.attributes['o:insetmode'] as string,
          },
          anchor: '',
          editAs: '',
          protection: {},
        };
        break;
      default:
        this.parser = this.map![node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag: {
        const textboxModel = this.map!['v:textbox'].model as Record<string, unknown> | undefined;
        const clientDataModel = this.map!['x:ClientData'].model as Record<string, unknown> | undefined;
        this.model.margins.inset = textboxModel?.inset as string | undefined;
        this.model.protection = (clientDataModel?.protection as Record<string, unknown>) || {};
        this.model.anchor = (clientDataModel?.anchor as string) || '';
        this.model.editAs = (clientDataModel?.editAs as string) || '';
        return false;
      }
      default:
        return true;
    }
  }
}

VmlShapeXform.V_SHAPE_ATTRIBUTES = (model, index) => ({
  id: `_x0000_s${1025 + index}`,
  type: '#_x0000_t202',
  style:
    'position:absolute; margin-left:105.3pt;margin-top:10.5pt;width:97.8pt;height:59.1pt;z-index:1;visibility:hidden',
  fillcolor: 'infoBackground [80]',
  strokecolor: 'none [81]',
  'o:insetmode': model.note.margins && model.note.margins.insetmode,
});

export default VmlShapeXform;