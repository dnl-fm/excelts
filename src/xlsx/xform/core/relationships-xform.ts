

import XmlStream from '../../../utils/xml-stream.ts';
import BaseXform from '../base-xform.ts';
import RelationshipXform from './relationship-xform.ts';

const RELATIONSHIPS_ATTRIBUTES = {
  xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships',
};

class RelationshipsXform extends BaseXform {
  declare model: unknown[] | undefined;
  declare map: Record<string, BaseXform>;

  constructor() {
    super();

    this.map = {
      Relationship: new RelationshipXform(),
    };
  }

  render(xmlStream: XmlStream, model?: unknown[]) {
    model = model || (this as any)._values;
    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode('Relationships', RELATIONSHIPS_ATTRIBUTES);

    (model || []).forEach(relationship => {
      this.map.Relationship.render(xmlStream, relationship);
    });

    xmlStream.closeNode();
  }

  parseOpen(node: { name: string; attributes?: Record<string, string> }) {
    if (this.parser) {
      this.parser.parseOpen?.(node);
      return true;
    }
    switch (node.name) {
      case 'Relationships':
        this.model = [];
        return true;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen?.(node);
          return true;
        }
        throw new Error(`Unexpected xml node in parseOpen: ${JSON.stringify(node)}`);
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
      }
      return true;
    }
    switch (name) {
      case 'Relationships':
        return false;
      default:
        throw new Error(`Unexpected xml node in parseClose: ${name}`);
    }
  }
}

export default RelationshipsXform;