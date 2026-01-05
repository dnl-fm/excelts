

import fs from 'fs';
import JSZip from 'jszip';
import ZipStream from '../utils/zip-stream.ts';
import StreamBuf from '../utils/stream-buf.ts';
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
import { PassThrough } from 'readable-stream';
import { bufferToString } from '../utils/browser-buffer-decode.ts';
import _RelType from './rel-type.ts';
import type {
  ReadableStream,
  XlsxReadOptions,
  XlsxWriteOptions,
  FsReadFileOptions,
  XlsxModel,
  ZipEntry,
  ZipWriter,
} from '../types/index.ts';

function fsReadFileAsync(filename: string, options?: FsReadFileOptions): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    fs.readFile(filename, options, (error, data) => {
      if (error) {
        reject(error);
      } else {
        resolve(data);
      }
    });
  });
}

/**
 * XLSX handles reading and writing XLSX containers for a workbook instance.
 */
class XLSX {
  workbook: unknown;

  constructor(workbook: unknown) {
    this.workbook = workbook;
  }

  // ===============================================================================
  // Workbook
  // =========================================================================
  // Read

  async readFile(filename: string, options?: XlsxReadOptions): Promise<unknown> {
    if (!(await utils.fs.exists(filename))) {
      throw new Error(`File not found: ${filename}`);
    }
    const stream = fs.createReadStream(filename);
    try {
      const workbook = await this.read(stream, options);
      stream.close();
      return workbook;
    } catch (error) {
      stream.close();
      throw error;
    }
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
            hyperlinks.hyperlink = ((drawingOptions.rels as Record<string, unknown>)[hyperlinks.rId!] as Record<string, unknown>).Target;
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
    const worksheet = await xform.parseStream(stream);
    (worksheet as Record<string, unknown>).sheetNo = sheetNo;
    if (path) {
      model.worksheetHash[path] = worksheet as any;
    }
    model.worksheets.push(worksheet as any);
  }

  async _processCommentEntry(stream: ReadableStream, model: XlsxModel, name: string): Promise<void> {
    const xform = new CommentsXform();
    const comments = await xform.parseStream(stream);
    model.comments[`../${name}.xml`] = comments as any;
  }

  async _processTableEntry(stream: ReadableStream, model: XlsxModel, name: string): Promise<void> {
    const xform = new TableXform();
    const table = await xform.parseStream(stream);
    model.tables[`../tables/${name}.xml`] = table as any;
  }

  async _processWorksheetRelsEntry(stream: ReadableStream, model: XlsxModel, sheetNo: number): Promise<void> {
    const xform = new RelationshipsXform();
    const relationships = await xform.parseStream(stream);
    model.worksheetRels[sheetNo] = relationships as any;
  }

  async _processMediaEntry(entry: ZipEntry, model: XlsxModel, filename: string): Promise<void> {
    const lastDot = filename.lastIndexOf('.');
    // if we can't determine extension, ignore it
    if (lastDot >= 1) {
      const extension = filename.substr(lastDot + 1);
      const name = filename.substr(0, lastDot);
      await new Promise<void>((resolve, reject) => {
        const streamBuf = new StreamBuf();
        streamBuf.on('finish', () => {
          model.mediaIndex[filename] = model.media.length;
          model.mediaIndex[name] = model.media.length;
          const medium = {
            type: 'image',
            name,
            extension,
            buffer: streamBuf.toBuffer(),
          };
          model.media.push(medium);
          resolve();
        });
        (entry as any).on('error', (error: Error) => {
          reject(error);
        });
        (entry as any).pipe(streamBuf);
      });
    }
  }

  async _processDrawingEntry(entry: ZipEntry, model: XlsxModel, name: string): Promise<void> {
    const xform = new DrawingXform();
    const drawing = await xform.parseStream(entry as any);
    model.drawings[name] = drawing as any;
  }

  async _processDrawingRelsEntry(entry: ZipEntry, model: XlsxModel, name: string): Promise<void> {
    const xform = new RelationshipsXform();
    const relationships = await xform.parseStream(entry as any);
    model.drawingRels[name] = relationships as any;
  }

  async _processVmlDrawingEntry(entry: ZipEntry, model: XlsxModel, name: string): Promise<void> {
    const xform = new VmlNotesXform();
    const vmlDrawing = await xform.parseStream(entry as any);
    model.vmlDrawings[`../drawings/${name}.vml`] = vmlDrawing as any;
  }

  async _processThemeEntry(entry: ZipEntry, model: XlsxModel, name: string): Promise<void> {
    await new Promise<void>((resolve, reject) => {
      // TODO: stream entry into buffer and store the xml in the model.themes[]
      const stream = new StreamBuf();
      (entry as any).on('error', reject);
      stream.on('error', reject);
      stream.on('finish', () => {
        model.themes[name] = stream.read().toString();
        resolve();
      });
      (entry as any).pipe(stream);
    });
  }

  async _processRawXmlEntry(entry: ZipEntry, bucket: Record<string, string>, name: string): Promise<void> {
    await new Promise<void>((resolve, reject) => {
      const stream = new StreamBuf();
      (entry as any).on('error', reject);
      stream.on('error', reject);
      stream.on('finish', () => {
        bucket[name] = stream.read().toString();
        resolve();
      });
      (entry as any).pipe(stream);
    });
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
    // TODO: Remove once node v8 is deprecated
    // Detect and upgrade old streams
    let readStream = stream;
    if (typeof (stream as any)[Symbol.asyncIterator] !== 'function' && typeof (stream as any).pipe === 'function') {
      readStream = (stream as any).pipe(new PassThrough());
    }
    const chunks: Buffer[] = [];
    for await (const chunk of readStream) {
      chunks.push(chunk as Buffer);
    }
    return this.load(Buffer.concat(chunks), options);
  }

  async load(data: Buffer, options?: XlsxReadOptions): Promise<unknown> {
    let buffer: Buffer;
    if (options && options.base64) {
      buffer = Buffer.from(data.toString(), 'base64');
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

    const zip = await JSZip.loadAsync(buffer);
    for (const entry of Object.values(zip.files)) {
      /* eslint-disable no-await-in-loop */
      if (!entry.dir) {
        let entryName = entry.name;
        if (entryName[0] === '/') {
          entryName = entryName.substr(1);
        }
        let stream: PassThrough;
        if (
          entryName.match(/xl\/media\//) ||
          entryName.match(/xl\/theme\/([a-zA-Z0-9]+)[.]xml/)
        ) {
          stream = new PassThrough();
          stream.write(await entry.async('nodebuffer'));
        } else {
          stream = new PassThrough({
            writableObjectMode: true,
            readableObjectMode: true,
          });
          let content: string;
          if ((process as any).browser) {
            content = bufferToString(await entry.async('nodebuffer'));
          } else {
            content = await entry.async('string');
          }
          const chunkSize = 16 * 1024;
          for (let i = 0; i < content.length; i += chunkSize) {
            stream.write(content.substring(i, i + chunkSize));
          }
        }
        stream.end();
        switch (entryName) {
          case '_rels/.rels':
            model.globalRels = await this.parseRels(stream);
            break;

          case 'xl/workbook.xml': {
            const workbook = await this.parseWorkbook(stream);
            model.sheets = (workbook as Record<string, unknown>).sheets as any;
            model.definedNames = (workbook as Record<string, unknown>).definedNames;
            model.views = (workbook as Record<string, unknown>).views as any;
            model.properties = (workbook as Record<string, unknown>).properties as any;
            model.calcProperties = (workbook as Record<string, unknown>).calcProperties;
            break;
          }

          case 'xl/_rels/workbook.xml.rels':
            model.workbookRels = await this.parseRels(stream);
            break;

          case 'xl/sharedStrings.xml':
            model.sharedStrings = new SharedStringsXform();
            await model.sharedStrings.parseStream(stream);
            break;

          case 'xl/styles.xml':
            model.styles = new StylesXform();
            await model.styles.parseStream(stream);
            break;

          case 'docProps/app.xml': {
            const appXform = new AppXform();
            const appProperties = await appXform.parseStream(stream);
            model.company = (appProperties as Record<string, unknown>).company as any;
            model.manager = (appProperties as Record<string, unknown>).manager as any;
            break;
          }

          case 'docProps/core.xml': {
            const coreXform = new CoreXform();
            const coreProperties = await coreXform.parseStream(stream);
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
              await this._processThemeEntry(entry as any, model, match[1]);
              break;
            }
            match = entryName.match(/xl\/media\/([a-zA-Z0-9]+[.][a-zA-Z0-9]{3,4})$/);
            if (match) {
              await this._processMediaEntry(entry as any, model, match[1]);
              break;
            }
            match = entryName.match(/xl\/drawings\/([a-zA-Z0-9]+)[.]xml/);
            if (match) {
              await this._processDrawingEntry(entry as any, model, match[1]);
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
              await this._processRawXmlEntry(entry as any, model.pivotTablesRaw, match[1]);
              break;
            }
            match = entryName.match(/xl\/pivotCache\/(pivotCacheDefinition\d+)[.]xml/);
            if (match) {
              await this._processRawXmlEntry(entry as any, model.pivotCacheDefinitionsRaw, match[1]);
              break;
            }
            match = entryName.match(/xl\/pivotCache\/(pivotCacheRecords\d+)[.]xml/);
            if (match) {
              await this._processRawXmlEntry(entry as any, model.pivotCacheRecordsRaw, match[1]);
              break;
            }
            match = entryName.match(/xl\/drawings\/_rels\/([a-zA-Z0-9]+)[.]xml[.]rels/);
            if (match) {
              await this._processDrawingRelsEntry(entry as any, model, match[1]);
              break;
            }
            match = entryName.match(/xl\/drawings\/(vmlDrawing\d+)[.]vml/);
            if (match) {
              await this._processVmlDrawingEntry(entry as any, model, match[1]);
              break;
            }
          }
        }
      }
    }

    this.reconcile(model, options);

    // apply model
    (this.workbook as any).model = model;
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
      const drawing = (worksheet as Record<string, unknown>).drawing;
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
      const tables = (worksheet as Record<string, unknown>).tables as any[];
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
        Type: (XLSX as any).RelType.PivotCacheRecords,
        Target: 'pivotCacheRecords1.xml',
      },
    ]);
    zip.append(xml, {name: 'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels'});

    xml = pivotTableXform.toXml(pivotTable);
    zip.append(xml, {name: 'xl/pivotTables/pivotTable1.xml'});

    xml = relsXform.toXml([
      {
        Id: 'rId1',
        Type: (XLSX as any).RelType.PivotCacheDefinition,
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
      {Id: 'rId1', Type: (XLSX as any).RelType.OfficeDocument, Target: 'xl/workbook.xml'},
      {Id: 'rId2', Type: (XLSX as any).RelType.CoreProperties, Target: 'docProps/core.xml'},
      {Id: 'rId3', Type: (XLSX as any).RelType.ExtenderProperties, Target: 'docProps/app.xml'},
    ]);
    zip.append(xml, {name: '_rels/.rels'});
  }

  async addWorkbookRels(zip: ZipWriter, model: XlsxModel): Promise<void> {
    let count = 1;
    const relationships: Record<string, unknown>[] = [
      {Id: `rId${count++}`, Type: (XLSX as any).RelType.Styles, Target: 'styles.xml'},
      {Id: `rId${count++}`, Type: (XLSX as any).RelType.Theme, Target: 'theme/theme1.xml'},
    ];
    if ((model.sharedStrings as Record<string, unknown>)?.count) {
      relationships.push({
        Id: `rId${count++}`,
        Type: (XLSX as any).RelType.SharedStrings,
        Target: 'sharedStrings.xml',
      });
    }
    if ((model.pivotTables || []).length) {
      const pivotTable = model.pivotTables[0];
      (pivotTable as Record<string, unknown>).rId = `rId${count++}`;
      relationships.push({
        Id: (pivotTable as Record<string, unknown>).rId,
        Type: (XLSX as any).RelType.PivotCacheDefinition,
        Target: 'pivotCache/pivotCacheDefinition1.xml',
      });
    }
    model.worksheets.forEach(worksheet => {
      (worksheet as Record<string, unknown>).rId = `rId${count++}`;
      relationships.push({
        Id: (worksheet as Record<string, unknown>).rId,
        Type: (XLSX as any).RelType.Worksheet,
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

      const rels = (worksheet as Record<string, unknown>).rels as any[];
      if (rels && rels.length) {
        xmlStream = new XmlStream();
        relationshipsXform.render(xmlStream, rels);
        zip.append(xmlStream.xml, {name: `xl/worksheets/_rels/sheet${(worksheet as Record<string, unknown>).id}.xml.rels`});
      }

      const comments = (worksheet as Record<string, unknown>).comments as any[];
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

    model.sharedStrings = new SharedStringsXform();
    model.styles = model.useStyles ? new StylesXform(true) : (new StylesXform as any).Mock();

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
    worksheetOptions.drawings = model.drawings = {};
    worksheetOptions.commentRefs = model.commentRefs = [];
    let tableCount = 0;
    model.tables = {};
    model.worksheets.forEach((worksheet: any) => {
      (worksheet.tables || []).forEach((table: any) => {
        tableCount++;
        table.target = `table${tableCount}.xml`;
        table.id = tableCount;
      });

      worksheetXform.prepare(worksheet, worksheetOptions);
    });
  }

  async write(stream: WritableStream, options?: XlsxWriteOptions): Promise<XLSX> {
    const opts = options || {};
    const model = ((this.workbook as any).model || {}) as XlsxModel;
    const zip = new (ZipStream as any).ZipWriter(opts.zip);
    zip.pipe(stream);

    this.prepareModel(model, opts);

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

  writeFile(filename: string, options?: XlsxWriteOptions): Promise<void> {
    const stream = fs.createWriteStream(filename);

    return new Promise((resolve, reject) => {
      stream.on('finish', () => {
        resolve();
      });
      stream.on('error', (error: Error) => {
        reject(error);
      });

      this.write(stream, options)
        .then(() => {
          stream.end();
        })
        .catch((err: Error) => {
          reject(err);
        });
    });
  }

  async writeBuffer(options?: XlsxWriteOptions): Promise<Buffer> {
    const stream = new StreamBuf();
    await this.write(stream as any, options);
    return stream.read();
  }
}

XLSX.RelType = _RelType;

export default XLSX;