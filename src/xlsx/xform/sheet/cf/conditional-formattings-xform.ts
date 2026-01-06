

import BaseXform from '../../base-xform.ts';
import ConditionalFormattingXform from './conditional-formatting-xform.ts';

class ConditionalFormattingsXform extends BaseXform {
  cfXform: ConditionalFormattingXform;
  declare model: unknown[] | undefined;

  constructor() {
    super();

    this.cfXform = new ConditionalFormattingXform();
  }

  get tag() {
    return 'conditionalFormatting';
  }

  reset() {
    this.model = [];
  }

  prepare(model: any[], options: any) {
    // ensure each rule has a priority value
    let nextPriority = model.reduce(
      (p, cf) => Math.max(p, ...cf.rules.map((rule: any) => rule.priority || 0)),
      1
    );
    model.forEach(cf => {
      cf.rules.forEach((rule: any) => {
        if (!rule.priority) {
          rule.priority = nextPriority++;
        }

        if (rule.style) {
          rule.dxfId = options.styles.addDxfStyle(rule.style);
        }
      });
    });
  }

  render(xmlStream: any, model: any[]) {
    model.forEach(cf => {
      this.cfXform.render(xmlStream, cf);
    });
  }

  parseOpen(node: { name: string }) {
    if (this.parser) {
      this.parser.parseOpen?.(node);
      return true;
    }

    switch (node.name) {
      case 'conditionalFormatting':
        this.parser = this.cfXform;
        this.parser.parseOpen?.(node);
        return true;

      default:
        return false;
    }
  }

  parseText(text: string) {
    if (this.parser) {
      this.parser.parseText?.(text);
    }
  }

  parseClose(name: string) {
    if (this.parser) {
      if (!this.parser.parseClose?.(name)) {
        (this.model as unknown[]).push(this.parser.model);
        this.parser = undefined;
        return false;
      }
      return true;
    }
    return false;
  }

  reconcile(model: any[], options: any) {
    model.forEach(cf => {
      cf.rules.forEach((rule: any) => {
        if (rule.dxfId !== undefined) {
          rule.style = options.styles.getDxfStyle(rule.dxfId);
          delete rule.dxfId;
        }
      });
    });
  }
}

export default ConditionalFormattingsXform;