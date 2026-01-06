
import BaseXform from '../base-xform.ts';
import type { SaxNode, XmlStreamWriter } from '../xform-types.ts';

interface DateXformOptions {
  tag: string;
  attr?: string;
  attrs?: Record<string, string>;
  format?: (dt: Date) => string;
  parse?: (str: string) => Date;
}

class DateXform extends BaseXform {
  declare tag: string;
  attr?: string;
  attrs?: Record<string, string>;
  _format: (dt: Date) => string;
  _parse: (str: string) => Date;
  text?: string[];

  constructor(options: DateXformOptions) {
    super();

    this.tag = options.tag;
    this.attr = options.attr;
    this.attrs = options.attrs;
    this._format =
      options.format ||
      function(dt: Date): string {
        try {
          if (Number.isNaN(dt.getTime())) return '';
          return dt.toISOString();
        } catch {
          return '';
        }
      };
    this._parse =
      options.parse ||
      function(str: string): Date {
        return new Date(str);
      };
  }

  render(xmlStream: XmlStreamWriter, model: Date): void {
    if (model) {
      xmlStream.openNode(this.tag);
      if (this.attrs) {
        xmlStream.addAttributes(this.attrs);
      }
      if (this.attr) {
        xmlStream.addAttribute(this.attr, this._format(model));
      } else {
        xmlStream.writeText(this._format(model));
      }
      xmlStream.closeNode();
    }
  }

  parseOpen(node: SaxNode): boolean {
    if (node.name === this.tag) {
      if (this.attr) {
        (this as unknown as { model: Date }).model = this._parse(node.attributes[this.attr] as string);
      } else {
        this.text = [];
      }
    }
    return true;
  }

  parseText(text: string): void {
    if (!this.attr && this.text) {
      this.text.push(text);
    }
  }

  parseClose(): boolean {
    if (!this.attr && this.text) {
      (this as unknown as { model: Date }).model = this._parse(this.text.join(''));
    }
    return false;
  }
}

export default DateXform;
