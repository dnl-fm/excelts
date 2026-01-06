
/**
 * Note wraps cell comments into the internal comment model.
 */
import _ from '../utils/under-dash.ts';

class Note {
  note: unknown;
  static DEFAULT_CONFIGS: Record<string, unknown>;

  constructor(note?: unknown) {
    this.note = note;
  }

  get model() {
    let value = null;
    switch (typeof this.note) {
      case 'string':
        value = {
          type: 'note',
          note: {
            texts: [
              {
                text: this.note,
              },
            ],
          },
        };
        break;
      default:
        value = {
          type: 'note',
          note: this.note,
        };
        break;
    }
    // Suitable for all cell comments
    return _.deepMerge({}, Note.DEFAULT_CONFIGS, value);
  }

  set model(value: Record<string, unknown>) {
    const note = value.note as { texts: Array<{ text?: string; [key: string]: unknown }> };
    const {texts} = note;
    if (texts.length === 1 && Object.keys(texts[0]).length === 1) {
      this.note = texts[0].text;
    } else {
      this.note = note;
    }
  }

  static fromModel(model: Record<string, unknown>): Note {
    const note = new Note();
    note.model = model;
    return note;
  }
}

Note.DEFAULT_CONFIGS = {
  note: {
    margins: {
      insetmode: 'auto',
      inset: [0.13, 0.13, 0.25, 0.25],
    },
    protection: {
      locked: 'True',
      lockText: 'True',
    },
    editAs: 'absolute',
  },
} as Record<string, unknown>;

export default Note;