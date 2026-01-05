
/* 'virtual' methods used as a form of documentation */
/* eslint-disable class-methods-use-this */

// Base class for Xforms
import parseSax from '../../utils/parse-sax.ts';
import XmlStream from '../../utils/xml-stream.ts';

class BaseXform {
  // constructor(/* model, name */) {}

  // ============================================================
  // Virtual Interface
  prepare(/* model, options */): void {
    // optional preparation (mutation) of model so it is ready for write
  }

  render(/* xmlStream, model */): void {
    // convert model to xml
  }

  parseOpen(_node: any): void {
    // XML node opened
  }

  parseText(_text: string): void {
    // chunk of text encountered for current node
  }

  parseClose(_name: string): boolean | undefined {
    // XML node closed
  }

  reconcile(_model: any, _options: any): void {
    // optional post-parse step (opposite to prepare)
  }

  // ============================================================
  reset(): void {
    // to make sure parses don't bleed to next iteration
    this.model = null;

    // if we have a map - reset them too
    if (this.map) {
      Object.values(this.map).forEach((xform: any) => {
        if (xform instanceof BaseXform) {
          xform.reset();
        } else if (xform.xform) {
          xform.xform.reset();
        }
      });
    }
  }

  mergeModel(obj: any): void {
    // set obj's props to this.model
    this.model = Object.assign(this.model || {}, obj);
  }

  async parse(saxParser: any): Promise<unknown> {
    for await (const events of saxParser) {
      for (const {eventType, value} of events) {
        if (eventType === 'opentag') {
          this.parseOpen(value);
        } else if (eventType === 'text') {
          this.parseText(value);
        } else if (eventType === 'closetag') {
          if (!this.parseClose(value.name)) {
            return this.model;
          }
        }
      }
    }
    return this.model;
  }

  async parseStream(stream: any): Promise<unknown> {
    return this.parse(parseSax(stream));
  }

  get xml(): string {
    // convenience function to get the xml of this.model
    // useful for manager types that are built during the prepare phase
    return this.toXml(this.model);
  }

  toXml(model: any): string {
    const xmlStream = new XmlStream();
    this.render(xmlStream, model);
    return xmlStream.xml;
  }

  // ============================================================
  // Useful Utilities
  static toAttribute(value: unknown, dflt?: unknown, always = false): string | undefined {
    if (value === undefined) {
      if (always) {
        return dflt as string | undefined;
      }
    } else if (always || value !== dflt) {
      return (value as any).toString();
    }
    return undefined;
  }

  static toStringAttribute(value: unknown, dflt?: unknown, always = false): string | undefined {
    return BaseXform.toAttribute(value, dflt, always);
  }

  static toStringValue(attr: string | undefined, dflt: unknown): unknown {
    return attr === undefined ? dflt : attr;
  }

  static toBoolAttribute(value: unknown, dflt?: unknown, always = false): string | undefined {
    if (value === undefined) {
      if (always) {
        return dflt as string | undefined;
      }
    } else if (always || value !== dflt) {
      return value ? '1' : '0';
    }
    return undefined;
  }

  static toBoolValue(attr: string | undefined, dflt: boolean): boolean {
    return attr === undefined ? dflt : attr === '1';
  }

  static toIntAttribute(value: unknown, dflt?: unknown, always = false): string | undefined {
    return BaseXform.toAttribute(value, dflt, always);
  }

  static toIntValue(attr: string | undefined, dflt: number): number {
    return attr === undefined ? dflt : parseInt(attr, 10);
  }

  static toFloatAttribute(value: unknown, dflt?: unknown, always = false): string | undefined {
    return BaseXform.toAttribute(value, dflt, always);
  }

  static toFloatValue(attr: string | undefined, dflt: number): number {
    return attr === undefined ? dflt : parseFloat(attr);
  }
}

export default BaseXform;