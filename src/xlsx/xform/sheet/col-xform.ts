
import utils from '../../../utils/utils.ts';
import BaseXform from '../base-xform.ts';
import type { SaxNode, XmlStreamWriter } from '../xform-types.ts';

interface ColModel {
  min: number;
  max: number;
  width?: number;
  styleId?: number;
  style?: unknown;
  hidden?: boolean;
  bestFit?: boolean;
  outlineLevel?: number;
  collapsed?: boolean;
  [key: string]: unknown;
}

interface ColPrepareOptions {
  styles: {
    addStyleModel: (style: unknown) => number | undefined;
  };
}

interface ColReconcileOptions {
  styles: {
    getStyleModel: (id: number) => unknown;
  };
}

class ColXform extends BaseXform {
  declare model: ColModel;

  get tag(): string {
    return 'col';
  }

  prepare(model: ColModel, options: ColPrepareOptions): void {
    const styleId = options.styles.addStyleModel(model.style || {});
    if (styleId) {
      model.styleId = styleId;
    }
  }

  render(xmlStream: XmlStreamWriter, model: ColModel): void {
    xmlStream.openNode('col');
    xmlStream.addAttribute('min', model.min);
    xmlStream.addAttribute('max', model.max);
    if (model.width) {
      xmlStream.addAttribute('width', model.width);
    }
    if (model.styleId) {
      xmlStream.addAttribute('style', model.styleId);
    }
    if (model.hidden) {
      xmlStream.addAttribute('hidden', '1');
    }
    if (model.bestFit) {
      xmlStream.addAttribute('bestFit', '1');
    }
    if (model.outlineLevel) {
      xmlStream.addAttribute('outlineLevel', model.outlineLevel);
    }
    if (model.collapsed) {
      xmlStream.addAttribute('collapsed', '1');
    }
    xmlStream.addAttribute('customWidth', '1');
    xmlStream.closeNode();
  }

  parseOpen(node: SaxNode): boolean {
    if (node.name === 'col') {
      const attrs = node.attributes as Record<string, string>;
      const model: ColModel = (this.model = {
        min: parseInt(attrs.min || '0', 10),
        max: parseInt(attrs.max || '0', 10),
        width:
          attrs.width === undefined
            ? undefined
            : parseFloat(attrs.width || '0'),
      });
      if (attrs.style) {
        model.styleId = parseInt(attrs.style, 10);
      }
      if (utils.parseBoolean(attrs.hidden)) {
        model.hidden = true;
      }
      if (utils.parseBoolean(attrs.bestFit)) {
        model.bestFit = true;
      }
      if (attrs.outlineLevel) {
        model.outlineLevel = parseInt(attrs.outlineLevel, 10);
      }
      if (utils.parseBoolean(attrs.collapsed)) {
        model.collapsed = true;
      }
      return true;
    }
    return false;
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }

  reconcile(model: ColModel, options: ColReconcileOptions): void {
    // reconcile column styles
    if (model.styleId) {
      model.style = options.styles.getStyleModel(model.styleId);
    }
  }
}

export default ColXform;