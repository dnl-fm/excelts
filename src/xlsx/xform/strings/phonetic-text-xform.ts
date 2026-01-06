

// <rPh sb="0" eb="1">
//   <t>(its pronounciation in KATAKANA)</t>
// </rPh>

import TextXform from './text-xform.ts';
import RichTextXform from './rich-text-xform.ts';
import BaseXform from '../base-xform.ts';

type PhoneticTextModel = {
  sb?: number;
  eb?: number;
  richText?: unknown[];
  text?: string;
};

class PhoneticTextXform extends BaseXform {
  declare model: PhoneticTextModel | undefined;

  constructor() {
    super();

    this.map = {
      r: new RichTextXform(),
      t: new TextXform(),
    };
  }

  get tag() {
    return 'rPh';
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      sb: model.sb || 0,
      eb: model.eb || 0,
    });
    if (model && model.hasOwnProperty('richText') && model.richText) {
      const {r} = this.map;
      model.richText.forEach(text => {
        r.render(xmlStream, text);
      });
    } else if (model) {
      this.map.t.render(xmlStream, model.text);
    }
    xmlStream.closeNode();
  }

  parseOpen(node) {
    const {name} = node;
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    if (name === this.tag) {
      this.model = {
        sb: parseInt(node.attributes.sb, 10),
        eb: parseInt(node.attributes.eb, 10),
      };
      return true;
    }
    this.parser = this.map[name];
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    return false;
  }

  parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string) {
    if (this.parser) {
      // Call child's parseClose - if it returns false, it's done
      const childDone = !this.parser.parseClose(name);
      if (childDone) {
        switch (name) {
          case 'r': {
            if (this.model) {
              let rt = this.model.richText;
              if (!rt) {
                rt = this.model.richText = [];
              }
              rt.push(this.parser.model);
            }
            break;
          }
          case 't':
            if (this.model) {
              this.model.text = this.parser.model as string;
            }
            break;
          default:
            break;
        }
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        return true;
    }
  }
}

export default PhoneticTextXform;