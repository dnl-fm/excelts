
import XmlStream from '../../../utils/xml-stream.ts';
import BaseXform from '../base-xform.ts';
import SharedStringXform from './shared-string-xform.ts';
import type { SaxNode, XmlStreamWriter } from '../xform-types.ts';

interface RichTextValue {
  richText: unknown[];
}

interface SharedStringsModel {
  values: unknown[];
  count: number;
  [key: string]: unknown;
}

class SharedStringsXform extends BaseXform {
  declare model: SharedStringsModel;
  hash: Record<string, number>;
  rich: Record<string, number>;
  _sharedStringXform?: SharedStringXform;

  constructor(model?: SharedStringsModel) {
    super();

    this.model = model || {
      values: [],
      count: 0,
    };
    this.hash = Object.create(null);
    this.rich = Object.create(null);
  }

  get sharedStringXform() {
    return this._sharedStringXform || (this._sharedStringXform = new SharedStringXform());
  }

  get values() {
    return this.model.values;
  }

  get uniqueCount() {
    return this.model.values.length;
  }

  get count() {
    return this.model.count;
  }

  getString(index: number): unknown {
    return this.model.values[index];
  }

  add(value: string | RichTextValue): number {
    return typeof value === 'object' && value !== null && 'richText' in value
      ? this.addRichText(value)
      : this.addText(value as string);
  }

  addText(value: string): number {
    let index = this.hash[value];
    if (index === undefined) {
      index = this.hash[value] = this.model.values.length;
      this.model.values.push(value);
    }
    this.model.count++;
    return index;
  }

  addRichText(value: RichTextValue): number {
    // TODO: add WeakMap here
    const xml = this.sharedStringXform.toXml(value);
    let index = this.rich[xml];
    if (index === undefined) {
      index = this.rich[xml] = this.model.values.length;
      this.model.values.push(value);
    }
    this.model.count++;
    return index;
  }

  // <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  // <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="<%=totalRefs%>" uniqueCount="<%=count%>">
  //   <si><t><%=text%></t></si>
  //   <si><r><rPr></rPr><t></t></r></si>
  // </sst>

  render(xmlStream: XmlStreamWriter, model?: SharedStringsModel): void {
    const m = model || this.model;
    xmlStream.openXml(XmlStream.StdDocAttributes);

    xmlStream.openNode('sst', {
      xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      count: m.count,
      uniqueCount: m.values.length,
    });

    const sx = this.sharedStringXform;
    m.values.forEach(sharedString => {
      sx.render(xmlStream, sharedString);
    });
    xmlStream.closeNode();
  }

  parseOpen(node: SaxNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'sst':
        return true;
      case 'si':
        this.parser = this.sharedStringXform;
        this.parser.parseOpen(node);
        return true;
      default:
        throw new Error(`Unexpected xml node in parseOpen: ${JSON.stringify(node)}`);
    }
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model.values.push(this.parser.model);
        this.model.count++;
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case 'sst':
        return false;
      default:
        throw new Error(`Unexpected xml node in parseClose: ${name}`);
    }
  }
}

export default SharedStringsXform;