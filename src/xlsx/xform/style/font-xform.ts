

// Font encapsulates translation from font model to xlsx
import ColorXform from './color-xform.ts';
import BooleanXform from '../simple/boolean-xform.ts';
import IntegerXform from '../simple/integer-xform.ts';
import StringXform from '../simple/string-xform.ts';
import UnderlineXform from './underline-xform.ts';
import _ from '../../../utils/under-dash.ts';
import BaseXform from '../base-xform.ts';
import type { SaxNode, XmlStreamWriter } from '../xform-types.ts';

interface FontXformOptions {
  tagName?: string;
  fontNameTag?: string;
}

interface FontMapEntry {
  prop: string;
  xform: BaseXform;
}

interface FontModel {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | string;
  charset?: number;
  color?: unknown;
  condense?: boolean;
  extend?: boolean;
  family?: number;
  outline?: boolean;
  vertAlign?: string;
  scheme?: string;
  shadow?: boolean;
  strike?: boolean;
  size?: number;
  name?: string;
  [key: string]: unknown;
}

class FontXform extends BaseXform {
  options: FontXformOptions;
  fontMap: Record<string, FontMapEntry>;
  static OPTIONS: FontXformOptions;

  constructor(options?: FontXformOptions) {
    super();

    this.options = options || FontXform.OPTIONS;

    this.fontMap = {
      b: {prop: 'bold', xform: new BooleanXform({tag: 'b', attr: 'val'})},
      i: {prop: 'italic', xform: new BooleanXform({tag: 'i', attr: 'val'})},
      u: {prop: 'underline', xform: new UnderlineXform()},
      charset: {prop: 'charset', xform: new IntegerXform({tag: 'charset', attr: 'val'})},
      color: {prop: 'color', xform: new ColorXform()},
      condense: {prop: 'condense', xform: new BooleanXform({tag: 'condense', attr: 'val'})},
      extend: {prop: 'extend', xform: new BooleanXform({tag: 'extend', attr: 'val'})},
      family: {prop: 'family', xform: new IntegerXform({tag: 'family', attr: 'val'})},
      outline: {prop: 'outline', xform: new BooleanXform({tag: 'outline', attr: 'val'})},
      vertAlign: {prop: 'vertAlign', xform: new StringXform({tag: 'vertAlign', attr: 'val'})},
      scheme: {prop: 'scheme', xform: new StringXform({tag: 'scheme', attr: 'val'})},
      shadow: {prop: 'shadow', xform: new BooleanXform({tag: 'shadow', attr: 'val'})},
      strike: {prop: 'strike', xform: new BooleanXform({tag: 'strike', attr: 'val'})},
      sz: {prop: 'size', xform: new IntegerXform({tag: 'sz', attr: 'val'})},
    };
    this.fontMap[this.options.fontNameTag!] = {
      prop: 'name',
      xform: new StringXform({tag: this.options.fontNameTag!, attr: 'val'}),
    };
  }

  get tag(): string | undefined {
    return this.options.tagName;
  }

  render(xmlStream: XmlStreamWriter, model: FontModel): void {
    const {fontMap} = this;

    xmlStream.openNode(this.options.tagName!);
    _.each(this.fontMap, (defn: FontMapEntry, tag: string) => {
      fontMap[tag].xform.render(xmlStream, model[defn.prop]);
    });
    xmlStream.closeNode();
  }

  parseOpen(node: SaxNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    if (this.fontMap[node.name]) {
      this.parser = this.fontMap[node.name].xform;
      return this.parser.parseOpen(node) as boolean;
    }
    switch (node.name) {
      case this.options.tagName:
        this.model = {};
        return true;
      default:
        return false;
    }
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this.parser && !this.parser.parseClose(name)) {
      const item = this.fontMap[name];
      if (this.parser.model) {
        (this.model as FontModel)[item.prop] = this.parser.model;
      }
      this.parser = undefined;
      return true;
    }
    switch (name) {
      case this.options.tagName:
        return false;
      default:
        return true;
    }
  }
}

FontXform.OPTIONS = {
  tagName: 'font',
  fontNameTag: 'name',
};

export default FontXform;