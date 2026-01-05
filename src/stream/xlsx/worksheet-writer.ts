

import _ from '../../utils/under-dash.ts';
import RelType from '../../xlsx/rel-type.ts';
import colCache from '../../utils/col-cache.ts';
import Encryptor from '../../utils/encryptor.ts';
import Dimensions from '../../doc/range.ts';
import StringBuf from '../../utils/string-buf.ts';
import Row from '../../doc/row.ts';
import Column from '../../doc/column.ts';
import SheetRelsWriter from './sheet-rels-writer.ts';
import SheetCommentsWriter from './sheet-comments-writer.ts';
import DataValidations from '../../doc/data-validations.ts';
import ListXform from '../../xlsx/xform/list-xform.ts';
import DataValidationsXform from '../../xlsx/xform/sheet/data-validations-xform.ts';
import SheetPropertiesXform from '../../xlsx/xform/sheet/sheet-properties-xform.ts';
import SheetFormatPropertiesXform from '../../xlsx/xform/sheet/sheet-format-properties-xform.ts';
import ColXform from '../../xlsx/xform/sheet/col-xform.ts';
import RowXform from '../../xlsx/xform/sheet/row-xform.ts';
import HyperlinkXform from '../../xlsx/xform/sheet/hyperlink-xform.ts';
import SheetViewXform from '../../xlsx/xform/sheet/sheet-view-xform.ts';
import SheetProtectionXform from '../../xlsx/xform/sheet/sheet-protection-xform.ts';
import PageMarginsXform from '../../xlsx/xform/sheet/page-margins-xform.ts';
import PageSetupXform from '../../xlsx/xform/sheet/page-setup-xform.ts';
import AutoFilterXform from '../../xlsx/xform/sheet/auto-filter-xform.ts';
import PictureXform from '../../xlsx/xform/sheet/picture-xform.ts';
import ConditionalFormattingsXform from '../../xlsx/xform/sheet/cf/conditional-formattings-xform.ts';
import HeaderFooterXform from '../../xlsx/xform/sheet/header-footer-xform.ts';
import RowBreaksXform from '../../xlsx/xform/sheet/row-breaks-xform.ts';
import type { WritableStream } from '../../types/index.ts';

export interface WorksheetWriterOptions {
  id: number;
  name?: string;
  state?: string;
  properties?: Record<string, unknown>;
  headerFooter?: Record<string, unknown>;
  pageSetup?: Record<string, unknown>;
  useSharedStrings?: boolean;
  workbook: unknown;
  views?: unknown[];
  autoFilter?: unknown;
}

export interface SheetProtectionOptions {
  spinCount?: number;
  [key: string]: unknown;
}

export interface BackgroundImage {
  imageId: string | number;
  rId?: string;
}

const xmlBuffer = new StringBuf();

// ============================================================================================
// Xforms

// since prepare and render are functional, we can use singletons
const xform = {
  dataValidations: new DataValidationsXform(),
  sheetProperties: new SheetPropertiesXform(),
  sheetFormatProperties: new SheetFormatPropertiesXform(),
  columns: new ListXform({tag: 'cols', length: false, childXform: new ColXform()}),
  row: new RowXform(),
  hyperlinks: new ListXform({tag: 'hyperlinks', length: false, childXform: new HyperlinkXform()}),
  sheetViews: new ListXform({tag: 'sheetViews', length: false, childXform: new SheetViewXform()}),
  sheetProtection: new SheetProtectionXform(),
  pageMargins: new PageMarginsXform(),
  pageSeteup: new PageSetupXform(),
  autoFilter: new AutoFilterXform(),
  picture: new PictureXform(),
  conditionalFormattings: new ConditionalFormattingsXform(),
  headerFooter: new HeaderFooterXform(),
  rowBreaks: new RowBreaksXform(),
};

// ============================================================================================

class WorksheetWriter {
  id: number;
  name: string;
  state: string;
  _rows: (Row | null)[];
  _columns: Column[] | null;
  _keys: Record<string, unknown>;
  _merges: Dimensions[];
  _sheetRelsWriter: SheetRelsWriter;
  _sheetCommentsWriter: SheetCommentsWriter;
  _dimensions: Dimensions;
  _rowZero: number;
  committed: boolean;
  dataValidations: DataValidations;
  _formulae: Record<string, unknown>;
  _siFormulae: number;
  conditionalFormatting: unknown[];
  rowBreaks: unknown[];
  properties: Record<string, unknown>;
  headerFooter: Record<string, unknown>;
  pageSetup: Record<string, unknown>;
  useSharedStrings: boolean;
  _workbook: unknown;
  hasComments: boolean;
  _views: unknown[];
  autoFilter: unknown;
  _media: unknown[];
  sheetProtection: Record<string, unknown> | null;
  startedData: boolean;
  _stream?: WritableStream;
  _background?: BackgroundImage;

  constructor(options: WorksheetWriterOptions) {
    this.id = options.id;
    this.name = options.name || `Sheet${this.id}`;
    this.state = options.state || 'visible';

    this._rows = [];
    this._columns = null;
    this._keys = {};

    const mergesArray = [] as any;
    mergesArray.add = function() {};
    this._merges = mergesArray;

    this._sheetRelsWriter = new SheetRelsWriter(options as any);
    this._sheetCommentsWriter = new SheetCommentsWriter(this, this._sheetRelsWriter, options as any);

    this._dimensions = new Dimensions();
    this._rowZero = 1;
    this.committed = false;

    this.dataValidations = new DataValidations();

    this._formulae = {};
    this._siFormulae = 0;

    this.conditionalFormatting = [];
    this.rowBreaks = [];

    this.properties = Object.assign(
      {},
      {
        defaultRowHeight: 15,
        dyDescent: 55,
        outlineLevelCol: 0,
        outlineLevelRow: 0,
      },
      options.properties
    );

    this.headerFooter = Object.assign(
      {},
      {
        differentFirst: false,
        differentOddEven: false,
        oddHeader: null,
        oddFooter: null,
        evenHeader: null,
        evenFooter: null,
        firstHeader: null,
        firstFooter: null,
      },
      options.headerFooter
    );

    this.pageSetup = Object.assign(
      {},
      {
        margins: {left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3},
        orientation: 'portrait',
        horizontalDpi: 4294967295,
        verticalDpi: 4294967295,
        fitToPage: !!(
          options.pageSetup &&
          ((options.pageSetup as Record<string, unknown>).fitToWidth || (options.pageSetup as Record<string, unknown>).fitToHeight) &&
          !(options.pageSetup as Record<string, unknown>).scale
        ),
        pageOrder: 'downThenOver',
        blackAndWhite: false,
        draft: false,
        cellComments: 'None',
        errors: 'displayed',
        scale: 100,
        fitToWidth: 1,
        fitToHeight: 1,
        paperSize: undefined,
        showRowColHeaders: false,
        showGridLines: false,
        horizontalCentered: false,
        verticalCentered: false,
        rowBreaks: null,
        colBreaks: null,
      },
      options.pageSetup
    );

    this.useSharedStrings = options.useSharedStrings || false;
    this._workbook = options.workbook;
    this.hasComments = false;

    this._views = options.views || [];
    this.autoFilter = options.autoFilter || null;
    this._media = [];

    this.sheetProtection = null;

    this._writeOpenWorksheet();

    this.startedData = false;
  }

  get workbook(): unknown {
    return this._workbook;
  }

  get stream(): WritableStream {
    if (!this._stream) {
      this._stream = (this._workbook as any)._openStream(`/xl/worksheets/sheet${this.id}.xml`);
      (this._stream as any).pause();
    }
    return this._stream;
  }

  // destroy - not a valid operation for a streaming writer
  // even though some streamers might be able to, it's a bad idea.
  destroy(): void {
    throw new Error('Invalid Operation: destroy');
  }

  commit(): void {
    if (this.committed) {
      return;
    }
    // commit all rows
    this._rows.forEach(cRow => {
      if (cRow) {
        // write the row to the stream
        this._writeRow(cRow);
      }
    });

    // we _cannot_ accept new rows from now on
    this._rows = null;

    if (!this.startedData) {
      this._writeOpenSheetData();
    }
    this._writeCloseSheetData();
    this._writeAutoFilter();
    this._writeMergeCells();

    // for some reason, Excel can't handle dimensions at the bottom of the file
    // this._writeDimensions();

    this._writeHyperlinks();
    this._writeConditionalFormatting();
    this._writeDataValidations();
    this._writeSheetProtection();
    this._writePageMargins();
    this._writePageSetup();
    this._writeBackground();
    this._writeHeaderFooter();
    this._writeRowBreaks();

    // Legacy Data tag for comments
    this._writeLegacyData();

    this._writeCloseWorksheet();
    // signal end of stream to workbook
    this.stream.end();

    this._sheetCommentsWriter.commit();
    // also commit the hyperlinks if any
    this._sheetRelsWriter.commit();

    this.committed = true;
  }

  get dimensions(): Dimensions {
    return this._dimensions;
  }

  get views(): unknown[] {
    return this._views;
  }

  get columns(): Column[] | null {
    return this._columns;
  }

  set columns(value: Column[]) {
    const headerCount = value.reduce((pv, cv) => {
      const count = ((cv as any).header && 1) || ((cv as any).headers?.length) || 0;
      return Math.max(pv, count);
    }, 0);
    (this as any)._headerRowCount = headerCount;

    let count = 1;
    const columns = (this._columns = []);
    value.forEach((defn: Column) => {
      const column = new Column(this, count++, false);
      columns.push(column);
      (column as any).defn = defn;
    });
  }

  getColumnKey(key: string): unknown {
    return this._keys[key];
  }

  setColumnKey(key: string, value: unknown): void {
    this._keys[key] = value;
  }

  deleteColumnKey(key: string): void {
    delete this._keys[key];
  }

  eachColumnKey(f: (column: unknown, key: string) => void): void {
    _.each(this._keys, f);
  }

  getColumn(c: number | string): Column {
    if (typeof c === 'string') {
      const col = this._keys[c];
      if (col) return col as Column;
      c = colCache.l2n(c);
    }
    if (!this._columns) {
      this._columns = [];
    }
    if (c > this._columns.length) {
      let n = this._columns.length + 1;
      while (n <= c) {
        this._columns.push(new Column(this, n++));
      }
    }
    return this._columns[c - 1];
  }

  get _nextRow(): number {
    return this._rowZero + this._rows.length;
  }

  eachRow(options: Record<string, unknown> | ((row: Row, num: number) => void), iteratee?: (row: Row, num: number) => void): void {
    let iter = iteratee;
    let opts: Record<string, unknown> | undefined = undefined;
    if (typeof options === 'function') {
      iter = options;
    } else {
      opts = options;
    }
    if (opts && opts.includeEmpty) {
      const n = this._nextRow;
      for (let i = this._rowZero; i < n; i++) {
        (iter as Function)(this.getRow(i), i);
      }
    } else {
      this._rows.forEach((row: Row | null) => {
        if (row && (row as any).hasValues) {
          (iter as Function)(row, (row as any).number);
        }
      });
    }
  }

  _commitRow(cRow: Row): void {
    let found = false;
    while (this._rows.length && !found) {
      const row = this._rows.shift();
      this._rowZero++;
      if (row) {
        this._writeRow(row);
        found = (row as any).number === (cRow as any).number;
        this._rowZero = (row as any).number + 1;
      }
    }
  }

  get lastRow(): Row | null | undefined {
    if (this._rows.length) {
      return this._rows[this._rows.length - 1];
    }
    return undefined;
  }

  findRow(rowNumber: number): Row | null | undefined {
    const index = rowNumber - this._rowZero;
    return this._rows[index];
  }

  getRow(rowNumber: number): Row {
    const index = rowNumber - this._rowZero;

    if (index < 0) {
      throw new Error('Out of bounds: this row has been committed');
    }
    let row = this._rows[index];
    if (!row) {
      this._rows[index] = row = new Row(this, rowNumber);
    }
    return row;
  }

  addRow(value: unknown): Row {
    const row = new Row(this, this._nextRow);
    this._rows[row.number - this._rowZero] = row;
    (row as any).values = value;
    return row;
  }

  findCell(r: number | string, c?: number): unknown {
    const address = colCache.getAddress(r, c);
    const row = this.findRow(address.row);
    return row ? (row as any).findCell(address.column) : undefined;
  }

  getCell(r: number | string, c?: number): unknown {
    const address = colCache.getAddress(r, c);
    const row = this.getRow(address.row);
    return (row as any).getCellEx(address);
  }

  mergeCells(...cells: unknown[]): void {
    const dimensions = new Dimensions(cells);

    this._merges.forEach((merge: Dimensions) => {
      if (merge.intersects(dimensions)) {
        throw new Error('Cannot merge already merged cells');
      }
    });

    const master = this.getCell(dimensions.top, dimensions.left);
    for (let i = dimensions.top; i <= dimensions.bottom; i++) {
      for (let j = dimensions.left; j <= dimensions.right; j++) {
        if (i > dimensions.top || j > dimensions.left) {
          const cell = this.getCell(i, j);
          (cell as any).merge(master);
        }
      }
    }

    this._merges.push(dimensions);
  }

  addConditionalFormatting(cf: unknown): void {
    this.conditionalFormatting.push(cf);
  }

  removeConditionalFormatting(filter: number | ((item: unknown) => boolean) | null): void {
    if (typeof filter === 'number') {
      this.conditionalFormatting.splice(filter, 1);
    } else if (typeof filter === 'function') {
      this.conditionalFormatting = this.conditionalFormatting.filter(filter);
    } else {
      this.conditionalFormatting = [];
    }
  }

  addBackgroundImage(imageId: string | number): void {
    this._background = {
      imageId,
    };
  }

  getBackgroundImageId(): string | number | undefined {
    return this._background?.imageId;
  }

  protect(password?: string, options?: SheetProtectionOptions): Promise<void> {
    return new Promise(resolve => {
      this.sheetProtection = {
        sheet: true,
      };
      if (options && 'spinCount' in options) {
        (options as any).spinCount = Number.isFinite((options as any).spinCount) ? Math.round(Math.max(0, (options as any).spinCount)) : 100000;
      }
      if (password) {
        (this.sheetProtection as any).algorithmName = 'SHA-512';
        (this.sheetProtection as any).saltValue = Encryptor.randomBytes(16).toString('base64');
        (this.sheetProtection as any).spinCount = options && 'spinCount' in options ? (options as any).spinCount : 100000;
        (this.sheetProtection as any).hashValue = Encryptor.convertPasswordToHash(
          password,
          'SHA512',
          (this.sheetProtection as any).saltValue,
          (this.sheetProtection as any).spinCount
        );
      }
      if (options) {
        this.sheetProtection = Object.assign(this.sheetProtection, options);
        if (!password && 'spinCount' in options) {
          delete (this.sheetProtection as any).spinCount;
        }
      }
      resolve();
    });
  }

  unprotect(): void {
    this.sheetProtection = null;
  }

  _write(text: string): void {
    xmlBuffer.reset();
    xmlBuffer.addText(text);
    (this.stream as any).write(xmlBuffer);
  }

  _writeSheetProperties(xmlBuf: StringBuf, properties: Record<string, unknown> | undefined, pageSetup: Record<string, unknown> | undefined): void {
    const sheetPropertiesModel: Record<string, unknown> = {
      outlineProperties: properties?.(properties as any).outlineProperties,
      tabColor: properties?.(properties as any).tabColor,
      pageSetup:
        pageSetup && (pageSetup as any).fitToPage
          ? {
              fitToPage: (pageSetup as any).fitToPage,
            }
          : undefined,
    };

    xmlBuf.addText((xform.sheetProperties as any).toXml(sheetPropertiesModel));
  }

  _writeSheetFormatProperties(xmlBuf: StringBuf, properties: Record<string, unknown> | undefined): void {
    const sheetFormatPropertiesModel: Record<string, unknown> = properties
      ? {
          defaultRowHeight: (properties as any).defaultRowHeight,
          dyDescent: (properties as any).dyDescent,
          outlineLevelCol: (properties as any).outlineLevelCol,
          outlineLevelRow: (properties as any).outlineLevelRow,
        }
      : undefined;
    if (properties && (properties as any).defaultColWidth) {
      (sheetFormatPropertiesModel as any).defaultColWidth = (properties as any).defaultColWidth;
    }

    xmlBuf.addText((xform.sheetFormatProperties as any).toXml(sheetFormatPropertiesModel));
  }

  _writeOpenWorksheet(): void {
    xmlBuffer.reset();

    xmlBuffer.addText('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    xmlBuffer.addText(
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
        ' mc:Ignorable="x14ac"' +
        ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
    );

    this._writeSheetProperties(xmlBuffer, this.properties, this.pageSetup);

    xmlBuffer.addText((xform.sheetViews as any).toXml(this._views));

    this._writeSheetFormatProperties(xmlBuffer, this.properties);

    (this.stream as any).write(xmlBuffer);
  }

  _writeColumns(): void {
    const cols = Column.toModel(this._columns);
    if (cols) {
      (xform.columns as any).prepare(cols, {styles: (this._workbook as any).styles});
      (this.stream as any).write((xform.columns as any).toXml(cols));
    }
  }

  _writeOpenSheetData(): void {
    this._write('<sheetData>');
  }

  _writeRow(row: Row): void {
    if (!this.startedData) {
      this._writeColumns();
      this._writeOpenSheetData();
      this.startedData = true;
    }

    if ((row as any).hasValues || (row as any).height) {
      const model = (row as any).model;
      const options: Record<string, unknown> = {
        styles: (this._workbook as any).styles,
        sharedStrings: this.useSharedStrings ? (this._workbook as any).sharedStrings : undefined,
        hyperlinks: (this._sheetRelsWriter as any).hyperlinksProxy,
        merges: this._merges,
        formulae: this._formulae,
        siFormulae: this._siFormulae,
        comments: [],
      };
      (xform.row as any).prepare(model, options);
      (this.stream as any).write((xform.row as any).toXml(model));

      if (((options.comments as any) || []).length) {
        this.hasComments = true;
        this._sheetCommentsWriter.addComments((options.comments as any));
      }
    }
  }

  _writeCloseSheetData(): void {
    this._write('</sheetData>');
  }

  _writeMergeCells(): void {
    if (this._merges.length) {
      xmlBuffer.reset();
      xmlBuffer.addText(`<mergeCells count="${this._merges.length}">`);
      this._merges.forEach((merge: Dimensions) => {
        xmlBuffer.addText(`<mergeCell ref="${merge}"/>`);
      });
      xmlBuffer.addText('</mergeCells>');

      (this.stream as any).write(xmlBuffer);
    }
  }

  _writeHyperlinks(): void {
    (this.stream as any).write((xform.hyperlinks as any).toXml((this._sheetRelsWriter as any)._hyperlinks));
  }

  _writeConditionalFormatting(): void {
    const options: Record<string, unknown> = {
      styles: (this._workbook as any).styles,
    };
    (xform.conditionalFormattings as any).prepare(this.conditionalFormatting, options);
    (this.stream as any).write((xform.conditionalFormattings as any).toXml(this.conditionalFormatting));
  }

  _writeRowBreaks(): void {
    (this.stream as any).write((xform.rowBreaks as any).toXml(this.rowBreaks));
  }

  _writeDataValidations(): void {
    (this.stream as any).write((xform.dataValidations as any).toXml((this.dataValidations as any).model));
  }

  _writeSheetProtection(): void {
    (this.stream as any).write((xform.sheetProtection as any).toXml(this.sheetProtection));
  }

  _writePageMargins(): void {
    (this.stream as any).write((xform.pageMargins as any).toXml((this.pageSetup as any).margins));
  }

  _writePageSetup(): void {
    (this.stream as any).write((xform.pageSeteup as any).toXml(this.pageSetup));
  }

  _writeHeaderFooter(): void {
    (this.stream as any).write((xform.headerFooter as any).toXml(this.headerFooter));
  }

  _writeAutoFilter(): void {
    (this.stream as any).write((xform.autoFilter as any).toXml(this.autoFilter));
  }

  _writeBackground(): void {
    if (this._background) {
      if (this._background.imageId !== undefined) {
        const image = (this._workbook as any).getImage(this._background.imageId);
        const pictureId = this._sheetRelsWriter.addMedia({
          Target: `../media/${(image as any).name}`,
          Type: RelType.Image,
        });

        this._background = {
          ...this._background,
          rId: pictureId,
        };
      }
      (this.stream as any).write((xform.picture as any).toXml({rId: this._background.rId}));
    }
  }

  _writeLegacyData(): void {
    if (this.hasComments) {
      xmlBuffer.reset();
      xmlBuffer.addText(`<legacyDrawing r:id="${(this._sheetCommentsWriter as any).vmlRelId}"/>`);
      (this.stream as any).write(xmlBuffer);
    }
  }

  _writeDimensions(): void {
    // for some reason, Excel can't handle dimensions at the bottom of the file
  }

  _writeCloseWorksheet(): void {
    this._write('</worksheet>');
  }
}

export default WorksheetWriter;