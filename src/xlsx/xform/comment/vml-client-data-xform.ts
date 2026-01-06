import BaseXform from '../base-xform.ts';
import VmlAnchorXform from './vml-anchor-xform.ts';
import VmlProtectionXform from './style/vml-protection-xform.ts';
import VmlPositionXform from './style/vml-position-xform.ts';
import type {XmlStreamWriter, SaxNode} from '../xform-types.ts';

const POSITION_TYPE = ['twoCells', 'oneCells', 'absolute'];

type VmlClientDataModel = {
  anchor: unknown[];
  protection: {
    locked?: unknown;
    lockText?: unknown;
  };
  editAs: string;
};

type VmlClientDataRenderModel = {
  note: {
    protection: {
      locked?: unknown;
      lockText?: unknown;
    };
    editAs?: string;
  };
  refAddress: {
    row: number;
    col: number;
  };
};

class VmlClientDataXform extends BaseXform {
  declare model: VmlClientDataModel | undefined;
  declare map: {
    'x:Anchor': VmlAnchorXform;
    'x:Locked': VmlProtectionXform;
    'x:LockText': VmlProtectionXform;
    'x:SizeWithCells': VmlPositionXform;
    'x:MoveWithCells': VmlPositionXform;
  };

  constructor() {
    super();
    this.map = {
      'x:Anchor': new VmlAnchorXform(),
      'x:Locked': new VmlProtectionXform({tag: 'x:Locked'}),
      'x:LockText': new VmlProtectionXform({tag: 'x:LockText'}),
      'x:SizeWithCells': new VmlPositionXform({tag: 'x:SizeWithCells'}),
      'x:MoveWithCells': new VmlPositionXform({tag: 'x:MoveWithCells'}),
    };
  }

  get tag() {
    return 'x:ClientData';
  }

  render(xmlStream: XmlStreamWriter, model: VmlClientDataRenderModel) {
    const {protection, editAs} = model.note;
    xmlStream.openNode(this.tag, {ObjectType: 'Note'});
    this.map['x:MoveWithCells'].render(xmlStream, editAs, POSITION_TYPE);
    this.map['x:SizeWithCells'].render(xmlStream, editAs, POSITION_TYPE);
    this.map['x:Anchor'].render(xmlStream, model);
    this.map['x:Locked'].render(xmlStream, protection.locked);
    xmlStream.leafNode('x:AutoFill', null, 'False');
    this.map['x:LockText'].render(xmlStream, protection.lockText);
    xmlStream.leafNode('x:Row', null, model.refAddress.row - 1);
    xmlStream.leafNode('x:Column', null, model.refAddress.col - 1);
    xmlStream.closeNode();
  }

  parseOpen(node: SaxNode) {
    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          anchor: [],
          protection: {},
          editAs: '',
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
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        this.normalizeModel();
        return false;
      default:
        return true;
    }
  }

  normalizeModel() {
    const position = Object.assign(
      {},
      this.map['x:MoveWithCells'].model as Record<string, unknown>,
      this.map['x:SizeWithCells'].model as Record<string, unknown>
    );
    const len = Object.keys(position).length;
    this.model!.editAs = POSITION_TYPE[len];
    this.model!.anchor = (this.map['x:Anchor'] as VmlAnchorXform & {text: unknown[]}).text;
    this.model!.protection.locked = (this.map['x:Locked'] as VmlProtectionXform & {text: unknown}).text;
    this.model!.protection.lockText = (this.map['x:LockText'] as VmlProtectionXform & {text: unknown}).text;
  }
}

export default VmlClientDataXform;
