

import BaseXform from '../base-xform.ts';
import utils from '../../../utils/utils.ts';
import CellXform from './cell-xform.ts';
import type {
  RowModel,
  RowXformPrepareOptions,
  RowXformReconcileOptions,
  XmlStreamWriter,
  SaxNode,
  CellModel,
} from '../xform-types.ts';

class RowXform extends BaseXform {
  maxItems?: number;
  numRowsSeen = 0;
  declare map: { c: CellXform; [key: string]: BaseXform };
  declare model: RowModel | undefined;

  constructor(options?: { maxItems?: number }) {
    super();

    this.maxItems = options && options.maxItems;
    this.map = {
      c: new CellXform(),
    };
  }

  get tag() {
    return 'row';
  }

  prepare(model: RowModel, options: RowXformPrepareOptions) {
    const styleId = options.styles.addStyleModel(model.style);
    if (styleId) {
      model.styleId = styleId;
    }
    const cellXform = this.map.c;
    model.cells.forEach(cellModel => {
      cellXform.prepare(cellModel, options);
    });
  }

  render(xmlStream: XmlStreamWriter, model: RowModel, options?: RowXformPrepareOptions) {
    xmlStream.openNode('row');
    xmlStream.addAttribute('r', model.number);
    if (model.height) {
      xmlStream.addAttribute('ht', model.height);
      xmlStream.addAttribute('customHeight', '1');
    }
    if (model.hidden) {
      xmlStream.addAttribute('hidden', '1');
    }
    if (model.min > 0 && model.max > 0 && model.min <= model.max) {
      xmlStream.addAttribute('spans', `${model.min}:${model.max}`);
    }
    if (model.styleId) {
      xmlStream.addAttribute('s', model.styleId);
      xmlStream.addAttribute('customFormat', '1');
    }
    xmlStream.addAttribute('x14ac:dyDescent', '0.25');
    if (model.outlineLevel) {
      xmlStream.addAttribute('outlineLevel', model.outlineLevel);
    }
    if (model.collapsed) {
      xmlStream.addAttribute('collapsed', '1');
    }

    const cellXform = this.map.c;
    model.cells.forEach(cellModel => {
      cellXform.render(xmlStream, cellModel, options);
    });

    xmlStream.closeNode();
  }

  parseOpen(node: SaxNode) {
    if (this.parser) {
      this.parser.parseOpen?.(node);
      return true;
    }
    if (node.name === 'row') {
      this.numRowsSeen += 1;
      const spans = node.attributes.spans
        ? node.attributes.spans.split(':').map(span => parseInt(span, 10))
        : [undefined, undefined];
      const model: RowModel = {
        number: parseInt(node.attributes.r, 10),
        min: spans[0],
        max: spans[1],
        cells: [],
      };
      this.model = model;
      if (node.attributes.s) {
        model.styleId = parseInt(node.attributes.s, 10);
      }
      if (utils.parseBoolean(node.attributes.hidden)) {
        model.hidden = true;
      }
      if (utils.parseBoolean(node.attributes.bestFit)) {
        model.bestFit = true;
      }
      if (node.attributes.ht) {
        model.height = parseFloat(node.attributes.ht);
      }
      if (node.attributes.outlineLevel) {
        model.outlineLevel = parseInt(node.attributes.outlineLevel, 10);
      }
      if (utils.parseBoolean(node.attributes.collapsed)) {
        model.collapsed = true;
      }
      return true;
    }

    this.parser = this.map[node.name];
    if (this.parser) {
      this.parser.parseOpen?.(node);
      return true;
    }
    return false;
  }

  parseText(text: string) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string) {
    if (this.parser) {
      if (!this.parser.parseClose?.(name)) {
        this.model!.cells.push(this.parser.model);
        if (this.maxItems && this.model!.cells.length > this.maxItems) {
          throw new Error(`Max column count (${this.maxItems}) exceeded`);
        }
        this.parser = undefined;
      }
      return true;
    }
    return false;
  }

  reconcile(model: RowModel, options: RowXformReconcileOptions) {
    model.style = model.styleId && options.styles ? options.styles.getStyleModel(model.styleId) : {};
    if (model.styleId !== undefined) {
      model.styleId = undefined;
    }

    const cellXform = this.map.c;
    model.cells.forEach(cellModel => {
      cellXform.reconcile(cellModel as CellModel, options);
    });
  }
}

export default RowXform;