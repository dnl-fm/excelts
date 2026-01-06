import BaseXform from '../../../../src/xlsx/xform/base-xform.ts';

interface ChildXform {
  name?: string;
  tag?: string;
  xform: {
    prepare: (model: unknown, options: unknown) => void;
    render: (xmlStream: unknown, model: unknown) => void;
    parseOpen: (node: {name: string}) => boolean;
    parseText: (text: string) => void;
    parseClose: (name: string) => boolean;
    model: unknown;
  };
}

interface CompyXformOptions {
  tag: string;
  attrs?: Record<string, unknown>;
  children: ChildXform[];
}

class CompyXform extends BaseXform {
  declare tag: string;
  declare attrs?: Record<string, unknown>;
  children!: ChildXform[];
  _map!: Record<string, ChildXform>;
  _parser?: ChildXform;

  constructor(options: CompyXformOptions) {
    super();

    this.tag = options.tag;
    this.attrs = options.attrs;
    this.children = options.children;
    this._map = this.children.reduce((map: Record<string, ChildXform>, child) => {
      const name = child.name || child.tag || '';
      const tag = child.tag || child.name || '';
      map[tag] = child;
      child.name = name;
      child.tag = tag;
      return map;
    }, {});
  }

  prepare(model: Record<string, unknown>, options: unknown) {
    this.children.forEach(child => {
      child.xform.prepare(model[child.tag!], options);
    });
  }

  render(xmlStream: {openNode: (tag: string, attrs?: unknown) => void; closeNode: () => void}, model: Record<string, unknown>) {
    xmlStream.openNode(this.tag, this.attrs);
    this.children.forEach(child => {
      child.xform.render(xmlStream, model[child.name!]);
    });
    xmlStream.closeNode();
  }

  parseOpen(node: {name: string}) {
    if (this._parser) {
      this._parser.xform.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.model = {};
        return true;
      default:
        this._parser = this._map[node.name];
        if (this._parser) {
          this._parser.xform.parseOpen(node);
          return true;
        }
    }
    return false;
  }

  parseText(text: string) {
    if (this._parser) {
      this._parser.xform.parseText(text);
    }
  }

  parseClose(name: string) {
    if (this._parser) {
      if (!this._parser.xform.parseClose(name)) {
        (this.model as Record<string, unknown>)[this._parser.name!] = this._parser.xform.model;
        this._parser = undefined;
      }
      return true;
    }
    return false;
  }

  reconcile(model: Record<string, unknown>, options: unknown) {
    this.children.forEach(child => {
      child.xform.prepare(model[child.tag!], options);
    });
  }
}

export default CompyXform;