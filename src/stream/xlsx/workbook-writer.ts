/**
 * WorkbookWriter streams XLSX output with optional shared strings and styles.
 */
import { Writable, PassThrough } from 'stream';
import { zipSync, strToU8 } from 'fflate';
import StreamBuf from '../../utils/stream-buf.ts';
import RelType from '../../xlsx/rel-type.ts';
import StylesXform from '../../xlsx/xform/style/styles-xform.ts';
import SharedStrings from '../../utils/shared-strings.ts';
import DefinedNames from '../../doc/defined-names.ts';
import CoreXform from '../../xlsx/xform/core/core-xform.ts';
import RelationshipsXform from '../../xlsx/xform/core/relationships-xform.ts';
import ContentTypesXform from '../../xlsx/xform/core/content-types-xform.ts';
import AppXform from '../../xlsx/xform/core/app-xform.ts';
import WorkbookXform from '../../xlsx/xform/book/workbook-xform.ts';
import SharedStringsXform from '../../xlsx/xform/strings/shared-strings-xform.ts';
import WorksheetWriter from './worksheet-writer.ts';
import theme1Xml from '../../xlsx/xml/theme1.ts';

/** Options for WorkbookWriter */
export interface WorkbookWriterOptions {
  stream?: Writable;
  filename?: string;
  useSharedStrings?: boolean;
  useStyles?: boolean;
  zip?: Record<string, unknown>;
  created?: Date;
  modified?: Date;
  creator?: string;
  lastModifiedBy?: string;
  lastPrinted?: Date;
}

/** Options for adding a worksheet */
export interface WorksheetWriterOptions {
  useSharedStrings?: boolean;
  properties?: Record<string, unknown>;
  state?: 'visible' | 'hidden' | 'veryHidden';
  pageSetup?: Record<string, unknown>;
  views?: unknown[];
  autoFilter?: unknown;
  headerFooter?: Record<string, unknown>;
  tabColor?: unknown;
}

/** Image data for WorkbookWriter */
export interface WorkbookWriterImage {
  buffer?: Buffer | ArrayBuffer;
  base64?: string;
  filename?: string;
  extension: string;
}

/** Simple ZIP builder that accumulates files and zips on finalize */
class ZipBuilder {
  private files: Map<string, Uint8Array> = new Map();
  private streams: Map<string, StreamBuf> = new Map();
  private eventHandlers: Map<string, Array<(...args: unknown[]) => void>> = new Map();

  on(event: string, handler: (...args: unknown[]) => void): this {
    if (!this.eventHandlers.has(event)) {
      this.eventHandlers.set(event, []);
    }
    this.eventHandlers.get(event)!.push(handler);
    return this;
  }

  private emit(event: string, ...args: unknown[]): void {
    const handlers = this.eventHandlers.get(event);
    if (handlers) {
      handlers.forEach(h => h(...args));
    }
  }

  append(data: string | Buffer | Uint8Array | StreamBuf, options: { name: string; base64?: boolean }): void {
    // Normalize the path (remove leading slash)
    const name = options.name.startsWith('/') ? options.name.slice(1) : options.name;
    
    if (data instanceof StreamBuf) {
      // Track the stream - we'll read it on finalize
      this.streams.set(name, data);
      return;
    }

    let bytes: Uint8Array;
    
    if (options.base64) {
      // Decode base64 string to bytes
      const binaryString = atob(data as string);
      bytes = new Uint8Array(binaryString.length);
      for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
      }
    } else if (typeof data === 'string') {
      bytes = strToU8(data);
    } else if (Buffer.isBuffer(data)) {
      bytes = new Uint8Array(data);
    } else {
      bytes = data;
    }
    
    this.files.set(name, bytes);
  }

  async file(filename: string, options: { name: string }): Promise<void> {
    const name = options.name.startsWith('/') ? options.name.slice(1) : options.name;
    const data = await Bun.file(filename).arrayBuffer();
    this.files.set(name, new Uint8Array(data));
  }

  finalize(): Uint8Array {
    // Collect all files
    const filesObj: Record<string, Uint8Array> = {};
    
    // Add accumulated files
    for (const [name, data] of this.files) {
      filesObj[name] = data;
    }
    
    // Add stream contents
    for (const [name, stream] of this.streams) {
      const buffer = stream.toBuffer();
      if (buffer) {
        filesObj[name] = new Uint8Array(buffer);
      }
    }
    
    const zipped = zipSync(filesObj, { level: 6 });
    this.emit('finish');
    return zipped;
  }
}

class WorkbookWriter {
  created: Date;
  modified: Date;
  creator: string;
  lastModifiedBy: string;
  lastPrinted?: Date;
  useSharedStrings: boolean;
  sharedStrings: SharedStrings;
  styles: StylesXform;
  views: unknown[];
  media: unknown[];
  commentRefs: unknown[];
  zip: ZipBuilder;
  stream: Writable | StreamBuf;
  promise: Promise<unknown[]>;
  zipOptions?: Record<string, unknown>;
  private _definedNames: DefinedNames;
  private _worksheets: WorksheetWriter[];
  private _filename?: string;

  constructor(options?: WorkbookWriterOptions) {
    options = options || {};

    this.created = options.created || new Date();
    this.modified = options.modified || this.created;
    this.creator = options.creator || 'ExcelJS';
    this.lastModifiedBy = options.lastModifiedBy || 'ExcelJS';
    this.lastPrinted = options.lastPrinted;

    // using shared strings creates a smaller xlsx file but may use more memory
    this.useSharedStrings = options.useSharedStrings || false;
    this.sharedStrings = new SharedStrings();

    // style manager
    this.styles = options.useStyles ? new StylesXform(true) : new StylesXform.Mock(true);

    // defined names
    this._definedNames = new DefinedNames();

    this._worksheets = [];
    this.views = [];

    this.zipOptions = options.zip;

    this.media = [];
    this.commentRefs = [];

    this.zip = new ZipBuilder();
    if (options.stream) {
      this.stream = options.stream;
    } else if (options.filename) {
      this._filename = options.filename;
      this.stream = new StreamBuf();
    } else {
      this.stream = new StreamBuf();
    }

    // these bits can be added right now
    this.promise = Promise.all([this.addThemes(), this.addOfficeRels()]);
  }

  get definedNames(): DefinedNames {
    return this._definedNames;
  }

  _openStream(path: string): StreamBuf {
    const stream = new StreamBuf({bufSize: 65536, batch: true});
    this.zip.append(stream, {name: path});
    // Emit 'zipped' immediately since we're collecting data, not streaming
    stream.on('finish', () => {
      stream.emit('zipped');
    });
    return stream;
  }

  private _commitWorksheets(): Promise<void[]> | Promise<void> {
    const commitWorksheet = function(worksheet: WorksheetWriter): Promise<void> {
      if (!worksheet.committed) {
        return new Promise(resolve => {
          worksheet.stream.on('zipped', () => {
            resolve();
          });
          worksheet.commit();
        });
      }
      return Promise.resolve();
    };
    // if there are any uncommitted worksheets, commit them now and wait
    const promises = this._worksheets.map(commitWorksheet);
    if (promises.length) {
      return Promise.all(promises);
    }
    return Promise.resolve();
  }

  async commit(): Promise<WorkbookWriter> {
    // commit all worksheets, then add supplementary files
    await this.promise;
    await this.addMedia();
    await this._commitWorksheets();
    await Promise.all([
      this.addContentTypes(),
      this.addApp(),
      this.addCore(),
      this.addSharedStrings(),
      this.addStyles(),
      this.addWorkbookRels(),
    ]);
    await this.addWorkbook();
    return this._finalize();
  }

  get nextId(): number {
    // find the next unique spot to add worksheet
    let i;
    for (i = 1; i < this._worksheets.length; i++) {
      if (!this._worksheets[i]) {
        return i;
      }
    }
    return this._worksheets.length || 1;
  }

  addImage(image: WorkbookWriterImage): number {
    const id = this.media.length;
    const medium = Object.assign({}, image, {type: 'image', name: `image${id}.${image.extension}`});
    this.media.push(medium);
    return id;
  }

  getImage(id: number): unknown {
    return this.media[id];
  }

  addWorksheet(name?: string, options?: WorksheetWriterOptions): WorksheetWriter {
    // it's possible to add a worksheet with different than default
    // shared string handling
    // in fact, it's even possible to switch it mid-sheet
    options = options || {};
    const useSharedStrings =
      options.useSharedStrings !== undefined ? options.useSharedStrings : this.useSharedStrings;

    if (options.tabColor) {
      // eslint-disable-next-line no-console
      console.trace('tabColor option has moved to { properties: tabColor: {...} }');
      options.properties = Object.assign(
        {
          tabColor: options.tabColor,
        },
        options.properties
      );
    }

    const id = this.nextId;
    name = name || `sheet${id}`;

    const worksheet = new WorksheetWriter({
      id,
      name,
      workbook: this,
      useSharedStrings,
      properties: options.properties,
      state: options.state,
      pageSetup: options.pageSetup,
      views: options.views,
      autoFilter: options.autoFilter,
      headerFooter: options.headerFooter,
    });

    this._worksheets[id] = worksheet;
    return worksheet;
  }

  getWorksheet(id?: number | string): WorksheetWriter | undefined {
    if (id === undefined) {
      return this._worksheets.find(() => true);
    }
    if (typeof id === 'number') {
      return this._worksheets[id];
    }
    if (typeof id === 'string') {
      return this._worksheets.find(worksheet => worksheet && worksheet.name === id);
    }
    return undefined;
  }

  private addStyles(): Promise<void> {
    return new Promise(resolve => {
      this.zip.append(this.styles.xml, {name: 'xl/styles.xml'});
      resolve();
    });
  }

  private addThemes(): Promise<void> {
    return new Promise(resolve => {
      this.zip.append(theme1Xml, {name: 'xl/theme/theme1.xml'});
      resolve();
    });
  }

  private addOfficeRels(): Promise<void> {
    return new Promise(resolve => {
      const xform = new RelationshipsXform();
      const xml = xform.toXml([
        {Id: 'rId1', Type: RelType.OfficeDocument, Target: 'xl/workbook.xml'},
        {Id: 'rId2', Type: RelType.CoreProperties, Target: 'docProps/core.xml'},
        {Id: 'rId3', Type: RelType.ExtenderProperties, Target: 'docProps/app.xml'},
      ]);
      this.zip.append(xml, {name: '_rels/.rels'});
      resolve();
    });
  }

  private addContentTypes(): Promise<void> {
    return new Promise(resolve => {
      const model = {
        worksheets: this._worksheets.filter(Boolean),
        sharedStrings: this.sharedStrings,
        commentRefs: this.commentRefs,
        media: this.media,
      };
      const xform = new ContentTypesXform();
      const xml = xform.toXml(model);
      this.zip.append(xml, {name: '[Content_Types].xml'});
      resolve();
    });
  }

  private async addMedia(): Promise<void> {
    for (const medium of this.media as Array<{ type: string; name: string; filename?: string; buffer?: Buffer; base64?: string }>) {
      if (medium.type === 'image') {
        const filename = `xl/media/${medium.name}`;
        if (medium.filename) {
          await this.zip.file(medium.filename, {name: filename});
        } else if (medium.buffer) {
          this.zip.append(medium.buffer as Buffer, {name: filename});
        } else if (medium.base64) {
          const dataimg64 = medium.base64;
          const content = dataimg64.substring(dataimg64.indexOf(',') + 1);
          this.zip.append(content, {name: filename, base64: true});
        } else {
          throw new Error('Unsupported media');
        }
      } else {
        throw new Error('Unsupported media');
      }
    }
  }

  private addApp(): Promise<void> {
    return new Promise(resolve => {
      const model = {
        worksheets: this._worksheets.filter(Boolean),
      };
      const xform = new AppXform();
      const xml = xform.toXml(model);
      this.zip.append(xml, {name: 'docProps/app.xml'});
      resolve();
    });
  }

  private addCore(): Promise<void> {
    return new Promise(resolve => {
      const coreXform = new CoreXform();
      const xml = coreXform.toXml(this);
      this.zip.append(xml, {name: 'docProps/core.xml'});
      resolve();
    });
  }

  private addSharedStrings(): Promise<void> {
    if (this.sharedStrings.count) {
      return new Promise(resolve => {
        const sharedStringsXform = new SharedStringsXform();
        const xml = sharedStringsXform.toXml(this.sharedStrings);
        this.zip.append(xml, {name: 'xl/sharedStrings.xml'});
        resolve();
      });
    }
    return Promise.resolve();
  }

  private addWorkbookRels(): Promise<void> {
    let count = 1;
    const relationships = [
      {Id: `rId${count++}`, Type: RelType.Styles, Target: 'styles.xml'},
      {Id: `rId${count++}`, Type: RelType.Theme, Target: 'theme/theme1.xml'},
    ];
    if (this.sharedStrings.count) {
      relationships.push({
        Id: `rId${count++}`,
        Type: RelType.SharedStrings,
        Target: 'sharedStrings.xml',
      });
    }
    this._worksheets.forEach(worksheet => {
      if (worksheet) {
        worksheet.rId = `rId${count++}`;
        relationships.push({
          Id: worksheet.rId,
          Type: RelType.Worksheet,
          Target: `worksheets/sheet${worksheet.id}.xml`,
        });
      }
    });
    return new Promise(resolve => {
      const xform = new RelationshipsXform();
      const xml = xform.toXml(relationships);
      this.zip.append(xml, {name: 'xl/_rels/workbook.xml.rels'});
      resolve();
    });
  }

  private addWorkbook(): Promise<void> {
    const model = {
      worksheets: this._worksheets.filter(Boolean),
      definedNames: this._definedNames.model,
      views: this.views,
      properties: {},
      calcProperties: {},
    };

    return new Promise(resolve => {
      const xform = new WorkbookXform();
      xform.prepare(model);
      this.zip.append(xform.toXml(model), {name: 'xl/workbook.xml'});
      resolve();
    });
  }

  private async _finalize(): Promise<WorkbookWriter> {
    const zipped = this.zip.finalize();
    const buffer = Buffer.from(zipped);
    
    if (this._filename) {
      // Write directly to file with Bun
      await Bun.write(this._filename, buffer);
    } else if (this.stream instanceof StreamBuf) {
      // Write to StreamBuf
      this.stream.write(buffer);
      this.stream.end();
    } else {
      // Write to provided stream
      await new Promise<void>((resolve, reject) => {
        this.stream.write(buffer, (err: Error | null | undefined) => {
          if (err) reject(err);
          else {
            (this.stream as Writable).end(() => resolve());
          }
        });
      });
    }
    
    return this;
  }
}

export default WorkbookWriter;
