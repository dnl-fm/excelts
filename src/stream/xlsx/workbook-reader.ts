
import { unzipSync, strFromU8 } from 'fflate';
import { concat, isBytes, toBytes } from '../../utils/bytes.ts';
import parseSax from '../../utils/parse-sax.ts';
import StyleManager from '../../xlsx/xform/style/styles-xform.ts';
import WorkbookXform from '../../xlsx/xform/book/workbook-xform.ts';
import RelationshipsXform from '../../xlsx/xform/core/relationships-xform.ts';
import WorksheetReader from './worksheet-reader.ts';
import HyperlinkReader from './hyperlink-reader.ts';
import SimpleEventEmitter from '../../utils/event-emitter.ts';

/**
 * Create async iterable from string content
 */
function createAsyncIterable(content: string): AsyncIterable<string> {
  return {
    async *[Symbol.asyncIterator]() {
      const chunkSize = 64 * 1024;
      for (let i = 0; i < content.length; i += chunkSize) {
        yield content.substring(i, i + chunkSize);
      }
    }
  };
}

/** Options for WorkbookReader */
export interface WorkbookReaderOptions {
  worksheets?: 'emit' | 'ignore';
  sharedStrings?: 'cache' | 'emit' | 'ignore';
  hyperlinks?: 'cache' | 'emit' | 'ignore';
  styles?: 'cache' | 'ignore';
  entries?: 'emit' | 'ignore';
}

/** Event yielded during parsing */
export interface WorkbookReaderEvent {
  eventType: string;
  value: unknown;
}

/** Unzipped files from fflate */
type UnzippedFiles = Record<string, Uint8Array>;

/**
 * WorkbookReader streams XLSX contents and emits worksheet data.
 * Uses fflate for ZIP handling (Bun-native, no Node streams).
 */
class WorkbookReader extends SimpleEventEmitter {
  input: unknown;
  options: WorkbookReaderOptions;
  styles: StyleManager;
  sharedStrings?: unknown[];
  workbookRels?: unknown[];
  properties?: unknown;
  model?: unknown;
  private files?: UnzippedFiles;

  constructor(input: unknown, options: WorkbookReaderOptions = {}) {
    super();

    this.input = input;

    this.options = {
      worksheets: 'emit',
      sharedStrings: 'cache',
      hyperlinks: 'ignore',
      styles: 'ignore',
      entries: 'ignore',
      ...options,
    };

    this.styles = new StyleManager();
    this.styles.init();
  }

  private async _loadZip(input: unknown): Promise<UnzippedFiles> {
    let buffer: Uint8Array;
    
    if (typeof input === 'string') {
      // Filename - read with Bun
      const arrayBuffer = await Bun.file(input).arrayBuffer();
      buffer = new Uint8Array(arrayBuffer);
    } else if (input instanceof ArrayBuffer) {
      buffer = new Uint8Array(input);
    } else if (isBytes(input)) {
      buffer = input;
    } else if (input && typeof (input as AsyncIterable<Uint8Array>)[Symbol.asyncIterator] === 'function') {
      // Async iterable - collect chunks
      const chunks: Uint8Array[] = [];
      for await (const chunk of input as AsyncIterable<Uint8Array | string>) {
        chunks.push(isBytes(chunk) ? chunk : toBytes(chunk as string));
      }
      buffer = concat(chunks);
    } else {
      throw new Error(`Could not recognise input: ${input}`);
    }

    return unzipSync(buffer);
  }

  async read(input?: unknown, options?: WorkbookReaderOptions): Promise<void> {
    try {
      for await (const {eventType, value} of this.parse(input, options)) {
        switch (eventType) {
          case 'shared-strings':
            this.emit(eventType, value);
            break;
          case 'worksheet':
            this.emit(eventType, value);
            await (value as WorksheetReader).read();
            break;
          case 'hyperlinks':
            this.emit(eventType, value);
            break;
        }
      }
      this.emit('end');
      this.emit('finished');
    } catch (error) {
      this.emit('error', error);
    }
  }

  async *[Symbol.asyncIterator](): AsyncGenerator<unknown> {
    for await (const {eventType, value} of this.parse()) {
      if (eventType === 'worksheet') {
        yield value;
      }
    }
  }

  async *parse(input?: unknown, options?: WorkbookReaderOptions): AsyncGenerator<WorkbookReaderEvent> {
    if (options) this.options = options;
    
    // Load ZIP with fflate
    this.files = await this._loadZip(input || this.input);
    
    // Parse workbook relationships first
    const relsData = this.files['xl/_rels/workbook.xml.rels'];
    if (relsData) {
      const content = strFromU8(relsData);
      await this._parseRels(createAsyncIterable(content));
    }

    // Parse workbook
    const workbookData = this.files['xl/workbook.xml'];
    if (workbookData) {
      const content = strFromU8(workbookData);
      await this._parseWorkbook(createAsyncIterable(content));
    }

    // Parse shared strings
    const sharedStringsData = this.files['xl/sharedStrings.xml'];
    if (sharedStringsData) {
      const content = strFromU8(sharedStringsData);
      yield* this._parseSharedStrings(createAsyncIterable(content));
    }

    // Parse styles
    const stylesData = this.files['xl/styles.xml'];
    if (stylesData) {
      const content = strFromU8(stylesData);
      await this._parseStyles(createAsyncIterable(content));
    }

    // Find and parse worksheets
    const worksheetFiles = Object.keys(this.files)
      .filter(name => name.match(/xl\/worksheets\/sheet\d+\.xml$/))
      .sort((a, b) => {
        const numA = parseInt(a.match(/sheet(\d+)\.xml$/)?.[1] || '0');
        const numB = parseInt(b.match(/sheet(\d+)\.xml$/)?.[1] || '0');
        return numA - numB;
      });

    for (const path of worksheetFiles) {
      const match = path.match(/sheet(\d+)\.xml$/);
      if (match) {
        const sheetNo = match[1];
        const fileData = this.files[path];
        if (fileData) {
          const content = strFromU8(fileData);
          yield* this._parseWorksheet(createAsyncIterable(content), sheetNo);
        }

        // Check for hyperlinks
        const relsPath = `xl/worksheets/_rels/sheet${sheetNo}.xml.rels`;
        const relsData = this.files[relsPath];
        if (relsData) {
          const relsContent = strFromU8(relsData);
          yield* this._parseHyperlinks(createAsyncIterable(relsContent), sheetNo);
        }
      }
    }
  }

  private _emitEntry(payload: { type: string; id?: string }): void {
    if (this.options.entries === 'emit') {
      this.emit('entry', payload);
    }
  }

  private async _parseRels(stream: AsyncIterable<string>): Promise<void> {
    const xform = new RelationshipsXform();
    this.workbookRels = await xform.parseStream(stream);
  }

  private async _parseWorkbook(stream: AsyncIterable<string>): Promise<void> {
    this._emitEntry({type: 'workbook'});

    const workbook = new WorkbookXform();
    await workbook.parseStream(stream);

    this.properties = workbook.map.workbookPr;
    this.model = workbook.model;
  }

  private async *_parseSharedStrings(stream: AsyncIterable<string>): AsyncGenerator<WorkbookReaderEvent> {
    this._emitEntry({type: 'shared-strings'});
    switch (this.options.sharedStrings) {
      case 'cache':
        this.sharedStrings = [];
        break;
      case 'emit':
        break;
      default:
        return;
    }

    let text: string | null = null;
    let richText: Array<{font: unknown; text: string | null}> = [];
    let index = 0;
    let font: Record<string, unknown> | null = null;

    for await (const events of parseSax(stream)) {
      for (const {eventType, value} of events as Array<{eventType: string; value: unknown}>) {
        if (eventType === 'opentag') {
          const node = value as {name: string; attributes: Record<string, string>};
          switch (node.name) {
            case 'b':
              font = font || {};
              font.bold = true;
              break;
            case 'charset':
              font = font || {};
              font.charset = parseInt(node.attributes.charset, 10);
              break;
            case 'color':
              font = font || {};
              font.color = {};
              if (node.attributes.rgb) {
                (font.color as Record<string, unknown>).argb = node.attributes.rgb;
              }
              if (node.attributes.val) {
                (font.color as Record<string, unknown>).argb = node.attributes.val;
              }
              if (node.attributes.theme) {
                (font.color as Record<string, unknown>).theme = node.attributes.theme;
              }
              break;
            case 'family':
              font = font || {};
              font.family = parseInt(node.attributes.val, 10);
              break;
            case 'i':
              font = font || {};
              font.italic = true;
              break;
            case 'outline':
              font = font || {};
              font.outline = true;
              break;
            case 'rFont':
              font = font || {};
              font.name = node.attributes.val;
              break;
            case 'si':
              font = null;
              richText = [];
              text = null;
              break;
            case 'sz':
              font = font || {};
              font.size = parseInt(node.attributes.val, 10);
              break;
            case 'strike':
              break;
            case 't':
              text = null;
              break;
            case 'u':
              font = font || {};
              font.underline = true;
              break;
            case 'vertAlign':
              font = font || {};
              font.vertAlign = node.attributes.val;
              break;
          }
        } else if (eventType === 'text') {
          text = text ? text + (value as string) : (value as string);
        } else if (eventType === 'closetag') {
          const node = value as {name: string};
          switch (node.name) {
            case 'r':
              richText.push({
                font,
                text,
              });
              font = null;
              text = null;
              break;
            case 'si':
              if (this.options.sharedStrings === 'cache') {
                this.sharedStrings!.push(richText.length ? {richText} : text);
              } else if (this.options.sharedStrings === 'emit') {
                yield {eventType: 'shared-string', value: {index: index++, text: richText.length ? {richText} : text}};
              }
              richText = [];
              font = null;
              text = null;
              break;
          }
        }
      }
    }
  }

  private async _parseStyles(stream: AsyncIterable<string>): Promise<void> {
    this._emitEntry({type: 'styles'});
    if (this.options.styles === 'cache') {
      this.styles = new StyleManager();
      await this.styles.parseStream(stream);
    }
  }

  private *_parseWorksheet(stream: AsyncIterable<string>, sheetNo: string): Generator<WorkbookReaderEvent> {
    this._emitEntry({type: 'worksheet', id: sheetNo});
    const worksheetReader = new WorksheetReader({
      workbook: this,
      id: sheetNo,
      iterator: stream,
      options: this.options,
    });

    const matchingRel = (this.workbookRels as Array<{ Target: string; Id: string }> || [])
      .find(rel => rel.Target === `worksheets/sheet${sheetNo}.xml`);
    const matchingSheet = matchingRel && 
      ((this.model as { sheets?: Array<{ rId: string; id: number; name: string; state: string }> })?.sheets || [])
      .find(sheet => sheet.rId === matchingRel.Id);
    
    if (matchingSheet) {
      worksheetReader.id = matchingSheet.id;
      worksheetReader.name = matchingSheet.name;
      worksheetReader.state = matchingSheet.state;
    }
    
    if (this.options.worksheets === 'emit') {
      yield {eventType: 'worksheet', value: worksheetReader};
    }
  }

  private *_parseHyperlinks(stream: AsyncIterable<string>, sheetNo: string): Generator<WorkbookReaderEvent> {
    this._emitEntry({type: 'hyperlinks', id: sheetNo});
    const hyperlinksReader = new HyperlinkReader({
      workbook: this,
      id: sheetNo,
      iterator: stream,
      options: this.options,
    });
    if (this.options.hyperlinks === 'emit') {
      yield {eventType: 'hyperlinks', value: hyperlinksReader};
    }
  }
}

// for reference - these are the valid values for options
WorkbookReader.Options = {
  worksheets: ['emit', 'ignore'],
  sharedStrings: ['cache', 'emit', 'ignore'],
  hyperlinks: ['cache', 'emit', 'ignore'],
  styles: ['cache', 'ignore'],
  entries: ['emit', 'ignore'],
};

export default WorkbookReader;
