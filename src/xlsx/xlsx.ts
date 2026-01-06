
import { unzipSync, strFromU8 } from 'fflate';
import { concat, toBytes, isBytes, toString as bytesToString } from '../utils/bytes.ts';

/**
 * Create an async iterable from content (string or Uint8Array) for streaming parsers.
 * Yields content in chunks to simulate streaming behavior.
 */
function createAsyncIterable(content: string | Uint8Array, chunkSize = 16 * 1024): AsyncIterable<string | Uint8Array> {
  return {
    async *[Symbol.asyncIterator]() {
      if (typeof content === 'string') {
        for (let i = 0; i < content.length; i += chunkSize) {
          yield content.substring(i, i + chunkSize);
        }
      } else {
        for (let i = 0; i < content.length; i += chunkSize) {
          yield content.subarray(i, i + chunkSize);
        }
      }
    }
  };
}
import ZipStream from '../utils/zip-stream.ts';
import SimpleBuffer from '../utils/simple-buffer.ts';
import utils from '../utils/utils.ts';
import XmlStream from '../utils/xml-stream.ts';
import StylesXform from './xform/style/styles-xform.ts';
import CoreXform from './xform/core/core-xform.ts';
import SharedStringsXform from './xform/strings/shared-strings-xform.ts';
import RelationshipsXform from './xform/core/relationships-xform.ts';
import ContentTypesXform from './xform/core/content-types-xform.ts';
import AppXform from './xform/core/app-xform.ts';
import WorkbookXform from './xform/book/workbook-xform.ts';
import WorksheetXform from './xform/sheet/worksheet-xform.ts';
import DrawingXform from './xform/drawing/drawing-xform.ts';
import TableXform from './xform/table/table-xform.ts';
import PivotCacheRecordsXform from './xform/pivot-table/pivot-cache-records-xform.ts';
import PivotCacheDefinitionXform from './xform/pivot-table/pivot-cache-definition-xform.ts';
import PivotTableXform from './xform/pivot-table/pivot-table-xform.ts';
import CommentsXform from './xform/comment/comments-xform.ts';
import VmlNotesXform from './xform/comment/vml-notes-xform.ts';
import theme1Xml from './xml/theme1.ts';
import _RelType from './rel-type.ts';
import type Workbook from '../doc/workbook.ts';
import type { WorkbookModel } from '../doc/workbook.ts';
import type {
  ReadableStream,
  XlsxReadOptions,
  XlsxWriteOptions,
  FsReadFileOptions,
  XlsxModel,
  ZipWriter,
  ParsedWorksheetModel,
  ParsedCommentsModel,
  ParsedTableModel,
  ParsedRelationshipsModel,
  ParsedDrawingModel,
  ParsedVmlDrawingModel,
  ParsedWorkbookModel,
  ParsedAppPropertiesModel,
  ParsedCorePropertiesModel,
  RelationshipModel,
  CommentModel,
  TableModel,
  SharedStringsModel,
  WorksheetItemModel,
  DrawingModel,
} from '../types/index.ts';

async function fsReadFileAsync(filename: string, _options?: FsReadFileOptions): Promise<Uint8Array> {
  const arrayBuffer = await Bun.file(filename).arrayBuffer();
  return new Uint8Array(arrayBuffer);
}

/**
 * XLSX handles reading and writing XLSX containers for a workbook instance.
 */
class XLSX {
  static RelType: typeof _RelType;
  
  workbook: Workbook;

  constructor(workbook: Workbook) {
    this.workbook = workbook;
  }

  // ===============================================================================
  // Workbook
  // =========================================================================
  // Read

  async readFile(filename: string, options?: XlsxReadOptions): Promise<Workbook> {
    if (!(await utils.fs.exists(filename))) {
      throw new Error(`File not found: ${filename}`);
    }
    const arrayBuffer = await Bun.file(filename).arrayBuffer();
    const buffer = new Uint8Array(arrayBuffer);
    return this.load(buffer, options);
  }

  parseRels(stream: ReadableStream): Promise<unknown> {
    const xform = new RelationshipsXform();
    return xform.parseStream(stream);
  }

  parseWorkbook(stream: ReadableStream): Promise<unknown> {
    const xform = new WorkbookXform();
    return xform.parseStream(stream);
  }

  parseSharedStrings(stream: ReadableStream): Promise<unknown> {
    const xform = new SharedStringsXform();
    return xform.parseStream(stream);
  }

  reconcile(model: XlsxModel, options?: XlsxReadOptions): void {
    const workbookXform = new WorkbookXform();
    const worksheetXform = new WorksheetXform(options);
    const drawingXform = new DrawingXform();
    const tableXform = new TableXform();

    workbookXform.reconcile(model);

    // reconcile drawings with their rels
    const drawingOptions: Record<string, unknown> = {
      media: model.media,
      mediaIndex: model.mediaIndex,
    };
    Object.keys(model.drawings).forEach(name => {
      const drawing = model.drawings[name];
      const drawingRel = model.drawingRels[name];
      if (drawingRel) {
        drawingOptions.rels = drawingRel.reduce((o: Record<string, unknown>, rel) => {
          o[rel.Id] = rel;
          return o;
        }, {});
        (drawing.anchors || []).forEach(anchor => {
          const hyperlinks = anchor.picture && anchor.picture.hyperlinks;
          if (hyperlinks && (drawingOptions.rels as Record<string, unknown>)[hyperlinks.rId!]) {
            hyperlinks.hyperlink = ((drawingOptions.rels as Record<string, unknown>)[hyperlinks.rId!] as Record<string, unknown>).Target as string;
            delete hyperlinks.rId;
          }
        });
        drawingXform.reconcile(drawing, drawingOptions);
      }
    });

    // reconcile tables with the default styles
    const tableOptions: Record<string, unknown> = {
      styles: model.styles,
    };
    Object.values(model.tables).forEach(table => {
      tableXform.reconcile(table, tableOptions);
    });

    const sheetOptions: Record<string, unknown> = {
      styles: model.styles,
      sharedStrings: model.sharedStrings,
      media: model.media,
      mediaIndex: model.mediaIndex,
      date1904: model.properties && (model.properties as Record<string, unknown>).date1904,
      drawings: model.drawings,
      comments: model.comments,
      tables: model.tables,
      vmlDrawings: model.vmlDrawings,
    };
    model.worksheets.forEach(worksheet => {
      worksheet.relationships = model.worksheetRels[worksheet.sheetNo!];
      worksheetXform.reconcile(worksheet, sheetOptions);
    });

    // delete unnecessary parts
    delete model.worksheetHash;
    delete model.worksheetRels;
    delete model.globalRels;
    delete model.sharedStrings;
    delete model.workbookRels;
    delete model.styles;
    delete model.mediaIndex;
    delete model.drawings;
    delete model.drawingRels;
    delete model.vmlDrawings;
  }

  async _processWorksheetEntry(stream: ReadableStream, model: XlsxModel, sheetNo: number, options?: XlsxReadOptions, path?: string): Promise<void> {
    const xform = new WorksheetXform(options);
    const worksheet = (await xform.parseStream(stream)) as ParsedWorksheetModel;
    worksheet.sheetNo = sheetNo;
    if (path) {
      model.worksheetHash[path] = worksheet as unknown as WorksheetItemModel;
    }
    model.worksheets.push(worksheet as unknown as WorksheetItemModel);
  }

  async _processCommentEntry(stream: ReadableStream, model: XlsxModel, name: string): Promise<void> {
    const xform = new CommentsXform();
    const comments = (await xform.parseStream(stream)) as ParsedCommentsModel;
    model.comments[`../${name}.xml`] = comments;
  }

  async _processTableEntry(stream: ReadableStream, model: XlsxModel, name: string): Promise<void> {
    const xform = new TableXform();
    const table = (await xform.parseStream(stream)) as ParsedTableModel;
    model.tables[`../tables/${name}.xml`] = table;
  }

  async _processWorksheetRelsEntry(stream: ReadableStream, model: XlsxModel, sheetNo: number): Promise<void> {
    const xform = new RelationshipsXform();
    const relationships = (await xform.parseStream(stream)) as ParsedRelationshipsModel;
    model.worksheetRels[sheetNo] = relationships;
  }

  async _processMediaEntry(stream: ReadableStream, model: XlsxModel, filename: string): Promise<void> {
    const lastDot = filename.lastIndexOf('.');
    // if we can't determine extension, ignore it
    if (lastDot >= 1) {
      const extension = filename.substr(lastDot + 1);
      const name = filename.substr(0, lastDot);
      const chunks: Uint8Array[] = [];
      for await (const chunk of stream as AsyncIterable<Uint8Array>) {
        chunks.push(isBytes(chunk) ? chunk : toBytes(chunk as unknown as string));
      }
      const buffer = concat(chunks);
      model.mediaIndex[filename] = model.media.length;
      model.mediaIndex[name] = model.media.length;
      const medium = {
        type: 'image',
        name,
        extension,
        buffer,
      };
      model.media.push(medium);
    }
  }

  async _processDrawingEntry(stream: ReadableStream, model: XlsxModel, name: string): Promise<void> {
    const xform = new DrawingXform();
    const drawing = (await xform.parseStream(stream)) as ParsedDrawingModel;
    model.drawings[name] = drawing;
  }

  async _processDrawingRelsEntry(stream: ReadableStream, model: XlsxModel, name: string): Promise<void> {
    const xform = new RelationshipsXform();
    const relationships = (await xform.parseStream(stream)) as ParsedRelationshipsModel;
    model.drawingRels[name] = relationships;
  }

  async _processVmlDrawingEntry(stream: ReadableStream, model: XlsxModel, name: string): Promise<void> {
    const xform = new VmlNotesXform();
    const vmlDrawing = (await xform.parseStream(stream)) as ParsedVmlDrawingModel;
    model.vmlDrawings[`../drawings/${name}.vml`] = vmlDrawing;
  }

  async _processThemeEntry(stream: ReadableStream, model: XlsxModel, name: string): Promise<void> {
    const chunks: string[] = [];
    for await (const chunk of stream as AsyncIterable<Uint8Array | string>) {
      chunks.push(typeof chunk === 'string' ? chunk : bytesToString(chunk));
    }
    model.themes[name] = chunks.join('');
  }

  async _processRawXmlEntry(stream: ReadableStream, bucket: Record<string, string>, name: string): Promise<void> {
    const chunks: string[] = [];
    for await (const chunk of stream as AsyncIterable<Uint8Array | string>) {
      chunks.push(typeof chunk === 'string' ? chunk : bytesToString(chunk));
    }
    bucket[name] = chunks.join('');
  }

  /**
   * @deprecated since version 4.0. You should use `#read` instead. Please follow upgrade instruction: https://github.com/exceljs/exceljs/blob/master/UPGRADE-4.0.md
   */
  createInputStream() {
    throw new Error(
      '`XLSX#createInputStream` is deprecated. You should use `XLSX#read` instead. This method will be removed in version 5.0. Please follow upgrade instruction: https://github.com/exceljs/exceljs/blob/master/UPGRADE-4.0.md'
    );
  }

  async read(stream: ReadableStream, options?: XlsxReadOptions): Promise<unknown> {
    const streamAsObject = stream as unknown as Record<string | symbol, unknown>;
    if (typeof streamAsObject[Symbol.asyncIterator] !== 'function') {
      throw new Error('Stream must be async iterable. Use readFile() or load() instead.');
    }
    const chunks: Uint8Array[] = [];
    for await (const chunk of stream as AsyncIterable<Uint8Array>) {
      chunks.push(isBytes(chunk) ? chunk : toBytes(chunk as unknown as string));
    }
    return this.load(concat(chunks), options);
  }

  async load(data: Uint8Array, options?: XlsxReadOptions): Promise<Workbook> {
    let buffer: Uint8Array;
    if (options && options.base64) {
      buffer = toBytes(bytesToString(data), 'base64');
    } else {
      buffer = data;
    }

    const model: XlsxModel = {
      worksheets: [],
      worksheetHash: {},
      worksheetRels: {},
      themes: {},
      media: [],
      mediaIndex: {},
      drawings: {},
      drawingRels: {},
      comments: {},
      tables: {},
      vmlDrawings: {},
      pivotTablesRaw: {},
      pivotCacheDefinitionsRaw: {},
      pivotCacheRecordsRaw: {},
    };

    const unzipped = unzipSync(new Uint8Array(buffer));
    for (const [name, data] of Object.entries(unzipped)) {
      /* eslint-disable no-await-in-loop */
      let entryName = name;
      if (entryName[0] === '/') {
        entryName = entryName.substr(1);
      }
      let stream: AsyncIterable<string | Uint8Array>;
      if (
        entryName.match(/xl\/media\//) ||
        entryName.match(/xl\/theme\/([a-zA-Z0-9]+)[.]xml/)
      ) {
        stream = createAsyncIterable(new Uint8Array(data));
      } else {
        const content = strFromU8(data);
        stream = createAsyncIterable(content);
      }
      switch (entryName) {
          case '_rels/.rels':
            model.globalRels = (await this.parseRels(stream)) as RelationshipModel[];
            break;

          case 'xl/workbook.xml': {
            const workbook = (await this.parseWorkbook(stream)) as ParsedWorkbookModel;
            model.sheets = workbook.sheets;
            model.definedNames = workbook.definedNames;
            model.views = workbook.views;
            model.properties = workbook.properties;
            model.calcProperties = workbook.calcProperties;
            break;
          }

          case 'xl/_rels/workbook.xml.rels':
            model.workbookRels = (await this.parseRels(stream)) as RelationshipModel[];
            break;

          case 'xl/sharedStrings.xml':
            model.sharedStrings = new SharedStringsXform() as unknown as SharedStringsModel;
            await (model.sharedStrings as unknown as SharedStringsXform).parseStream(stream);
            break;

          case 'xl/styles.xml':
            model.styles = new StylesXform() as unknown as Record<string, unknown>;
            await (model.styles as unknown as StylesXform).parseStream(stream);
            break;

          case 'docProps/app.xml': {
            const appXform = new AppXform();
            const appProperties = (await appXform.parseStream(stream)) as ParsedAppPropertiesModel;
            model.company = appProperties.company;
            model.manager = appProperties.manager;
            break;
          }

          case 'docProps/core.xml': {
            const coreXform = new CoreXform();
            const coreProperties = (await coreXform.parseStream(stream)) as ParsedCorePropertiesModel;
            Object.assign(model, coreProperties);
            break;
          }

          default: {
            let match = entryName.match(/xl\/worksheets\/sheet(\d+)[.]xml/);
            if (match) {
              await this._processWorksheetEntry(stream, model, parseInt(match[1]), options, entryName);
              break;
            }
            match = entryName.match(/xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/);
            if (match) {
              await this._processWorksheetRelsEntry(stream, model, parseInt(match[1]));
              break;
            }
            match = entryName.match(/xl\/theme\/([a-zA-Z0-9]+)[.]xml/);
            if (match) {
              await this._processThemeEntry(stream, model, match[1]);
              break;
            }
            match = entryName.match(/xl\/media\/([a-zA-Z0-9]+[.][a-zA-Z0-9]{3,4})$/);
            if (match) {
              await this._processMediaEntry(stream, model, match[1]);
              break;
            }
            match = entryName.match(/xl\/drawings\/([a-zA-Z0-9]+)[.]xml/);
            if (match) {
              await this._processDrawingEntry(stream, model, match[1]);
              break;
            }
            match = entryName.match(/xl\/(comments\d+)[.]xml/);
            if (match) {
              await this._processCommentEntry(stream, model, match[1]);
              break;
            }
            match = entryName.match(/xl\/tables\/(table\d+)[.]xml/);
            if (match) {
              await this._processTableEntry(stream, model, match[1]);
              break;
            }
            match = entryName.match(/xl\/pivotTables\/(pivotTable\d+)[.]xml/);
            if (match) {
              await this._processRawXmlEntry(stream, model.pivotTablesRaw, match[1]);
              break;
            }
            match = entryName.match(/xl\/pivotCache\/(pivotCacheDefinition\d+)[.]xml/);
            if (match) {
              await this._processRawXmlEntry(stream, model.pivotCacheDefinitionsRaw, match[1]);
              break;
            }
            match = entryName.match(/xl\/pivotCache\/(pivotCacheRecords\d+)[.]xml/);
            if (match) {
              await this._processRawXmlEntry(stream, model.pivotCacheRecordsRaw, match[1]);
              break;
            }
            match = entryName.match(/xl\/drawings\/_rels\/([a-zA-Z0-9]+)[.]xml[.]rels/);
            if (match) {
              await this._processDrawingRelsEntry(stream, model, match[1]);
              break;
            }
            match = entryName.match(/xl\/drawings\/(vmlDrawing\d+)[.]vml/);
            if (match) {
              await this._processVmlDrawingEntry(stream, model, match[1]);
              break;
            }
          }
        }
    }

    this.reconcile(model, options);

    // apply model
    this.workbook.model = model as unknown as WorkbookModel;
    return this.workbook;
  }

  // =========================================================================
  // Write

  async addMedia(zip: ZipWriter, model: XlsxModel): Promise<void> {
    await Promise.all(
      model.media.map(async medium => {
        if (medium.type === 'image') {
          const filename = `xl/media/${medium.name}.${medium.extension}`;
          if (medium.filename) {
            const data = await fsReadFileAsync(medium.filename);
            return zip.append(data, {name: filename});
          }
          if (medium.buffer) {
            return zip.append(medium.buffer, {name: filename});
          }
          if (medium.base64) {
            const dataimg64 = medium.base64;
            const content = dataimg64.substring(dataimg64.indexOf(',') + 1);
            return zip.append(content, {name: filename, base64: true});
          }
        }
        throw new Error('Unsupported media');
      })
    );
  }

  addDrawings(zip: ZipWriter, model: XlsxModel): void {
    const drawingXform = new DrawingXform();
    const relsXform = new RelationshipsXform();

    model.worksheets.forEach(worksheet => {
      const drawing = (worksheet as Record<string, unknown>).drawing as DrawingModel | undefined;
      if (drawing) {
        drawingXform.prepare(drawing, {});
        let xml = drawingXform.toXml(drawing);
        zip.append(xml, {name: `xl/drawings/${(drawing as Record<string, unknown>).name}.xml`});

        xml = relsXform.toXml((drawing as Record<string, unknown>).rels);
        zip.append(xml, {name: `xl/drawings/_rels/${(drawing as Record<string, unknown>).name}.xml.rels`});
      }
    });
  }

  addTables(zip: ZipWriter, model: XlsxModel): void {
    const tableXform = new TableXform();

    model.worksheets.forEach(worksheet => {
      const tables = (worksheet as Record<string, unknown>).tables as TableModel[] | undefined;
      tables?.forEach(table => {
        tableXform.prepare(table, {});
        const tableXml = tableXform.toXml(table);
        zip.append(tableXml, {name: `xl/tables/${(table as Record<string, unknown>).target}`});
      });
    });
  }

  addPivotTables(zip: ZipWriter, model: XlsxModel): void {
    if (!(model.pivotTables?.length)) return;

    const pivotTable = model.pivotTables[0];

    const pivotCacheRecordsXform = new PivotCacheRecordsXform();
    const pivotCacheDefinitionXform = new PivotCacheDefinitionXform();
    const pivotTableXform = new PivotTableXform();
    const relsXform = new RelationshipsXform();

    let xml = pivotCacheRecordsXform.toXml(pivotTable);
    zip.append(xml, {name: 'xl/pivotCache/pivotCacheRecords1.xml'});

    xml = pivotCacheDefinitionXform.toXml(pivotTable);
    zip.append(xml, {name: 'xl/pivotCache/pivotCacheDefinition1.xml'});

    xml = relsXform.toXml([
      {
        Id: 'rId1',
        Type: _RelType.PivotCacheRecords,
        Target: 'pivotCacheRecords1.xml',
      },
    ]);
    zip.append(xml, {name: 'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels'});

    xml = pivotTableXform.toXml(pivotTable);
    zip.append(xml, {name: 'xl/pivotTables/pivotTable1.xml'});

    xml = relsXform.toXml([
      {
        Id: 'rId1',
        Type: _RelType.PivotCacheDefinition,
        Target: '../pivotCache/pivotCacheDefinition1.xml',
      },
    ]);
    zip.append(xml, {name: 'xl/pivotTables/_rels/pivotTable1.xml.rels'});
  }

  async addContentTypes(zip: ZipWriter, model: XlsxModel): Promise<void> {
    const xform = new ContentTypesXform();
    const xml = xform.toXml(model);
    zip.append(xml, {name: '[Content_Types].xml'});
  }

  async addApp(zip: ZipWriter, model: XlsxModel): Promise<void> {
    const xform = new AppXform();
    const xml = xform.toXml(model);
    zip.append(xml, {name: 'docProps/app.xml'});
  }

  async addCore(zip: ZipWriter, model: XlsxModel): Promise<void> {
    const coreXform = new CoreXform();
    zip.append(coreXform.toXml(model), {name: 'docProps/core.xml'});
  }

  async addThemes(zip: ZipWriter, model: XlsxModel): Promise<void> {
    const themes = model.themes || {theme1: theme1Xml};
    Object.keys(themes).forEach(name => {
      const xml = themes[name];
      const path = `xl/theme/${name}.xml`;
      zip.append(xml, {name: path});
    });
  }

  async addOfficeRels(zip: ZipWriter): Promise<void> {
    const xform = new RelationshipsXform();
    const xml = xform.toXml([
      {Id: 'rId1', Type: _RelType.OfficeDocument, Target: 'xl/workbook.xml'},
      {Id: 'rId2', Type: _RelType.CoreProperties, Target: 'docProps/core.xml'},
      {Id: 'rId3', Type: _RelType.ExtenderProperties, Target: 'docProps/app.xml'},
    ]);
    zip.append(xml, {name: '_rels/.rels'});
  }

  async addWorkbookRels(zip: ZipWriter, model: XlsxModel): Promise<void> {
    let count = 1;
    const relationships: Record<string, unknown>[] = [
      {Id: `rId${count++}`, Type: _RelType.Styles, Target: 'styles.xml'},
      {Id: `rId${count++}`, Type: _RelType.Theme, Target: 'theme/theme1.xml'},
    ];
    if ((model.sharedStrings as Record<string, unknown>)?.count) {
      relationships.push({
        Id: `rId${count++}`,
        Type: _RelType.SharedStrings,
        Target: 'sharedStrings.xml',
      });
    }
    if ((model.pivotTables || []).length) {
      const pivotTable = model.pivotTables[0];
      (pivotTable as Record<string, unknown>).rId = `rId${count++}`;
      relationships.push({
        Id: (pivotTable as Record<string, unknown>).rId,
        Type: _RelType.PivotCacheDefinition,
        Target: 'pivotCache/pivotCacheDefinition1.xml',
      });
    }
    model.worksheets.forEach(worksheet => {
      (worksheet as Record<string, unknown>).rId = `rId${count++}`;
      relationships.push({
        Id: (worksheet as Record<string, unknown>).rId,
        Type: _RelType.Worksheet,
        Target: `worksheets/sheet${(worksheet as Record<string, unknown>).id}.xml`,
      });
    });
    const xform = new RelationshipsXform();
    const xml = xform.toXml(relationships);
    zip.append(xml, {name: 'xl/_rels/workbook.xml.rels'});
  }

  async addSharedStrings(zip: ZipWriter, model: XlsxModel): Promise<void> {
    if (model.sharedStrings && (model.sharedStrings as Record<string, unknown>).count) {
      zip.append((model.sharedStrings as Record<string, unknown>).xml as string, {name: 'xl/sharedStrings.xml'});
    }
  }

  async addStyles(zip: ZipWriter, model: XlsxModel): Promise<void> {
    const xml = (model.styles as Record<string, unknown>)?.xml;
    if (xml) {
      zip.append(xml as string, {name: 'xl/styles.xml'});
    }
  }

  async addWorkbook(zip: ZipWriter, model: XlsxModel): Promise<void> {
    const xform = new WorkbookXform();
    zip.append(xform.toXml(model), {name: 'xl/workbook.xml'});
  }

  async addWorksheets(zip: ZipWriter, model: XlsxModel): Promise<void> {
    const worksheetXform = new WorksheetXform();
    const relationshipsXform = new RelationshipsXform();
    const commentsXform = new CommentsXform();
    const vmlNotesXform = new VmlNotesXform();

    model.worksheets.forEach(worksheet => {
      let xmlStream = new XmlStream();
      worksheetXform.render(xmlStream, worksheet);
      zip.append(xmlStream.xml, {name: `xl/worksheets/sheet${(worksheet as Record<string, unknown>).id}.xml`});

      const rels = (worksheet as Record<string, unknown>).rels as RelationshipModel[] | undefined;
      if (rels && rels.length) {
        xmlStream = new XmlStream();
        relationshipsXform.render(xmlStream, rels);
        zip.append(xmlStream.xml, {name: `xl/worksheets/_rels/sheet${(worksheet as Record<string, unknown>).id}.xml.rels`});
      }

      const comments = (worksheet as Record<string, unknown>).comments as CommentModel[] | undefined;
      if (comments && comments.length > 0) {
        xmlStream = new XmlStream();
        commentsXform.render(xmlStream, worksheet);
        zip.append(xmlStream.xml, {name: `xl/comments${(worksheet as Record<string, unknown>).id}.xml`});

        xmlStream = new XmlStream();
        vmlNotesXform.render(xmlStream, worksheet);
        zip.append(xmlStream.xml, {name: `xl/drawings/vmlDrawing${(worksheet as Record<string, unknown>).id}.vml`});
      }
    });
  }

  _finalize(zip: ZipWriter): Promise<XLSX> {
    return new Promise((resolve, reject) => {
      zip.on('finish', () => {
        resolve(this);
      });
      zip.on('error', (error: Error) => {
        reject(error);
      });
      zip.finalize();
    });
  }

  prepareModel(model: XlsxModel, options: XlsxWriteOptions): void {
    model.creator = model.creator || 'ExcelJS';
    model.lastModifiedBy = model.lastModifiedBy || 'ExcelJS';
    model.created = model.created || new Date();
    model.modified = model.modified || new Date();

    model.useSharedStrings = options.useSharedStrings !== undefined ? options.useSharedStrings : true;
    model.useStyles = options.useStyles !== undefined ? options.useStyles : true;

    model.sharedStrings = new SharedStringsXform() as unknown as SharedStringsModel;
    const stylesXformRecord = StylesXform as unknown as Record<string, unknown>;
    const MockClass = stylesXformRecord.Mock as typeof StylesXform;
    model.styles = (model.useStyles ? new StylesXform(true) : new MockClass()) as unknown as Record<string, unknown>;

    const workbookXform = new WorkbookXform();
    const worksheetXform = new WorksheetXform();

    workbookXform.prepare(model);

    const worksheetOptions: Record<string, unknown> = {
      sharedStrings: model.sharedStrings,
      styles: model.styles,
      date1904: (model.properties as Record<string, unknown>)?.date1904,
      drawingsCount: 0,
      media: model.media,
    };
    worksheetOptions.drawings = [];
    model.drawings = {};
    worksheetOptions.commentRefs = model.commentRefs = [];
    let tableCount = 0;
    model.tables = {};
    model.worksheets.forEach((worksheet: unknown) => {
      const worksheetTables = (worksheet as Record<string, unknown>).tables as TableModel[] | undefined;
      (worksheetTables || []).forEach((table) => {
        tableCount++;
        table.target = `table${tableCount}.xml`;
        table.id = tableCount;
      });

      worksheetXform.prepare(worksheet, worksheetOptions);
    });
  }

  async write(stream: WritableStream, options?: XlsxWriteOptions): Promise<XLSX> {
    const opts = options || {};
    const model = ((this.workbook.model as unknown) || {}) as XlsxModel;
    const ZipWriterClass = ZipStream.ZipWriter;
    const zip = new ZipWriterClass(opts.zip as Record<string, unknown>) as ZipWriter;
    zip.pipe(stream);

    this.prepareModel(model, opts);

    // Collect tables and drawings from worksheets into arrays for content-types
    const allTables: TableModel[] = [];
    const allDrawings: Array<{ name: string }> = [];
    model.worksheets.forEach((worksheet: unknown) => {
      const ws = worksheet as Record<string, unknown>;
      const worksheetTables = ws.tables as TableModel[] | undefined;
      if (worksheetTables) {
        allTables.push(...worksheetTables);
      }
      const drawing = ws.drawing as { name: string } | undefined;
      if (drawing) {
        allDrawings.push(drawing);
      }
    });
    (model as unknown as Record<string, unknown>).tables = allTables;
    (model as unknown as Record<string, unknown>).drawings = allDrawings;

    await this.addContentTypes(zip, model);
    await this.addOfficeRels(zip);
    await this.addWorkbookRels(zip, model);
    await this.addWorksheets(zip, model);
    await this.addSharedStrings(zip, model);
    await this.addDrawings(zip, model);
    await this.addTables(zip, model);
    await this.addPivotTables(zip, model);
    await Promise.all([this.addThemes(zip, model), this.addStyles(zip, model)]);
    await this.addMedia(zip, model);
    await Promise.all([this.addApp(zip, model), this.addCore(zip, model)]);
    await this.addWorkbook(zip, model);
    return this._finalize(zip);
  }

  async writeFile(filename: string, options?: XlsxWriteOptions): Promise<void> {
    const buffer = await this.writeBuffer(options);
    await Bun.write(filename, buffer);
  }

  async writeBuffer(options?: XlsxWriteOptions): Promise<Uint8Array> {
    const buffer = new SimpleBuffer();
    await this.write(buffer as unknown as WritableStream, options);
    return buffer.read();
  }
}

XLSX.RelType = _RelType;

export default XLSX;
