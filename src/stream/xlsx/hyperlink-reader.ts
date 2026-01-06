

import parseSax from '../../utils/parse-sax.ts';
import Enums from '../../doc/enums.ts';
import RelType from '../../xlsx/rel-type.ts';
import SimpleEventEmitter from '../../utils/event-emitter.ts';

class HyperlinkReader extends SimpleEventEmitter {
  workbook: unknown;
  id: string;
  iterator: AsyncIterable<string | Uint8Array>;
  options: Record<string, unknown>;
  hyperlinks?: Record<string, unknown>;

  constructor({workbook, id, iterator, options}: any) {
    super();

    this.workbook = workbook;
    this.id = id;
    this.iterator = iterator;
    this.options = options;
  }

  get count(): number {
    return (this.hyperlinks && Object.keys(this.hyperlinks).length) || 0;
  }

  each(fn: (value: unknown) => void): void {
    if (this.hyperlinks) {
      Object.values(this.hyperlinks).forEach(fn);
    }
  }

  async read(): Promise<void> {
    const {iterator, options} = this;
    let emitHyperlinks = false;
    let hyperlinks: Record<string, unknown> | null = null;
    switch (options.hyperlinks) {
      case 'emit':
        emitHyperlinks = true;
        break;
      case 'cache':
        this.hyperlinks = hyperlinks = {};
        break;
      default:
        break;
    }

    if (!emitHyperlinks && !hyperlinks) {
      this.emit('finished');
      return;
    }

    try {
      for await (const events of parseSax(iterator)) {
        for (const {eventType, value} of events) {
          if (eventType === 'opentag') {
            const node = value;
            if (node.name === 'Relationship') {
              const rId = node.attributes.Id;
              switch (node.attributes.Type) {
                case RelType.Hyperlink:
                  {
                    const relationship = {
                      type: Enums.RelationshipType.Styles,
                      rId,
                      target: node.attributes.Target,
                      targetMode: node.attributes.TargetMode,
                    };
                    if (emitHyperlinks) {
                      this.emit('hyperlink', relationship);
                    } else {
                      hyperlinks[relationship.rId] = relationship;
                    }
                  }
                  break;

                default:
                  break;
              }
            }
          }
        }
      }
      this.emit('finished');
    } catch (error) {
      this.emit('error', error);
    }
  }
}

export default HyperlinkReader;