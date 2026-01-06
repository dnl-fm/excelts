
import BaseXform from './base-xform.ts';

interface ListXformOptions {
  tag: string;
  always?: boolean;
  count?: boolean;
  empty?: boolean;
  $count?: string;
  $?: Record<string, unknown>;
  childXform: BaseXform;
  maxItems?: number;
  length?: number | boolean;
}

class ListXform extends BaseXform {
  tag: string;
  always: boolean;
  count?: boolean;
  empty?: boolean;
  $count: string;
  $?: Record<string, unknown>;
  childXform: BaseXform;
  maxItems?: number;
  length?: number | boolean;

  constructor(options: ListXformOptions) {
    super();

    this.tag = options.tag;
    this.always = !!options.always;
    this.count = options.count;
    this.empty = options.empty;
    this.$count = options.$count || 'count';
    this.$ = options.$;
    this.childXform = options.childXform;
    this.maxItems = options.maxItems;
  }

  prepare(model: unknown[] | undefined, options: Record<string, unknown>) {
    const {childXform} = this;
    if (model) {
      model.forEach((childModel, index) => {
        options.index = index;
        childXform.prepare(childModel, options);
      });
    }
  }

  render(xmlStream: { openNode: Function; addAttribute: Function; closeNode: Function; leafNode: Function }, model: unknown[] | undefined) {
    if (this.always || (model && model.length)) {
      xmlStream.openNode(this.tag, this.$);
      if (this.count) {
        xmlStream.addAttribute(this.$count, (model && model.length) || 0);
      }

      const {childXform} = this;
      (model || []).forEach((childModel, index) => {
        childXform.render(xmlStream, childModel, index);
      });

      xmlStream.closeNode();
    } else if (this.empty) {
      xmlStream.leafNode(this.tag);
    }
  }

  parseOpen(node: { name: string }) {
    if (this.parser) {
      this.parser.parseOpen?.(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.model = [] as unknown as Record<string, unknown>;
        return true;
      default:
        this.childXform.parseOpen(node);
        if (this.childXform.model !== undefined) {
          this.parser = this.childXform;
          return true;
        }
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
      const result = this.parser.parseClose?.(name);
      if (!result) {
        (this.model as unknown as unknown[]).push(this.parser.model);
        this.parser = undefined;

        if (this.maxItems && (this.model as unknown as unknown[]).length > this.maxItems) {
          throw new Error(`Max ${(this.childXform as {tag?: string}).tag} count (${this.maxItems}) exceeded`);
        }
      }
      return true;
    }

    return false;
  }

  reconcile(model: unknown[] | undefined, options: unknown) {
    if (model) {
      const {childXform} = this;
      model.forEach(childModel => {
        childXform.reconcile(childModel, options);
      });
    }
  }
}

export default ListXform;