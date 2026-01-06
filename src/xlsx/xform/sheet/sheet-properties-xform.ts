
import BaseXform from '../base-xform.ts';
import ColorXform from '../style/color-xform.ts';
import PageSetupPropertiesXform from './page-setup-properties-xform.ts';
import OutlinePropertiesXform from './outline-properties-xform.ts';

type SheetPropertiesModel = {
  tabColor?: unknown;
  pageSetup?: unknown;
  outlineProperties?: unknown;
} | null;

class SheetPropertiesXform extends BaseXform {
  declare model: SheetPropertiesModel | undefined;

  constructor() {
    super();

    this.map = {
      tabColor: new ColorXform('tabColor'),
      pageSetUpPr: new PageSetupPropertiesXform(),
      outlinePr: new OutlinePropertiesXform(),
    };
  }

  get tag() {
    return 'sheetPr';
  }

  render(xmlStream: unknown, model: Record<string, unknown>): void {
    if (model) {
      const stream = xmlStream as { addRollback(): void; openNode(name: string): void; closeNode(): void; commit(): void; rollback(): void };
      stream.addRollback();
      stream.openNode('sheetPr');

      // These render methods return boolean indicating if they rendered content
      const tabColorRendered = (this.map!.tabColor as { render(s: unknown, m: unknown): boolean }).render(xmlStream, model.tabColor);
      const pageSetUpPrRendered = (this.map!.pageSetUpPr as { render(s: unknown, m: unknown): boolean }).render(xmlStream, model.pageSetup);
      const outlinePrRendered = (this.map!.outlinePr as { render(s: unknown, m: unknown): boolean }).render(xmlStream, model.outlineProperties);
      const inner = tabColorRendered || pageSetUpPrRendered || outlinePrRendered;

      if (inner) {
        stream.closeNode();
        stream.commit();
      } else {
        stream.rollback();
      }
    }
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    if (node.name === this.tag) {
      this.reset();
      return true;
    }
    if (this.map[node.name]) {
      this.parser = this.map[node.name];
      this.parser.parseOpen(node);
      return true;
    }
    return false;
  }

  parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
      return true;
    }
    return false;
  }

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    if (this.map.tabColor.model || this.map.pageSetUpPr.model || this.map.outlinePr.model) {
      this.model = {};
      if (this.map.tabColor.model) {
        this.model.tabColor = this.map.tabColor.model;
      }
      if (this.map.pageSetUpPr.model) {
        this.model.pageSetup = this.map.pageSetUpPr.model;
      }
      if (this.map.outlinePr.model) {
        this.model.outlineProperties = this.map.outlinePr.model;
      }
    } else {
      this.model = null;
    }
    return false;
  }
}

export default SheetPropertiesXform;