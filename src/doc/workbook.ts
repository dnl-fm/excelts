
import Worksheet from './worksheet.ts';
import DefinedNames from './defined-names.ts';
import XLSX from '../xlsx/xlsx.ts';
import CSV from '../csv/csv.ts';

/** Image data for adding to workbook */
export interface WorkbookImage {
  buffer?: Uint8Array | ArrayBuffer;
  base64?: string;
  filename?: string;
  extension: 'png' | 'jpeg' | 'gif';
}

/** Worksheet creation options */
export interface WorksheetOptions {
  properties?: {
    tabColor?: { argb?: string; theme?: number; indexed?: number };
    defaultRowHeight?: number;
    defaultColWidth?: number;
    dyDescent?: number;
  };
  pageSetup?: Record<string, unknown>;
  views?: unknown[];
  state?: 'visible' | 'hidden' | 'veryHidden';
  headerFooter?: Record<string, unknown>;
}

/** Workbook model for serialization */
export interface WorkbookModel {
  creator: string;
  lastModifiedBy: string;
  lastPrinted?: Date;
  created: Date;
  modified: Date;
  properties: Record<string, unknown>;
  worksheets: unknown[];
  sheets: unknown[];
  definedNames: unknown;
  views: unknown[];
  company: string;
  manager: string;
  title: string;
  subject: string;
  keywords: string;
  category: string;
  description: string;
  language?: string;
  revision?: number;
  contentStatus?: string;
  themes?: unknown;
  media: unknown[];
  pivotTables: unknown[];
  pivotTablesRaw: Record<string, unknown>;
  pivotCacheDefinitionsRaw: Record<string, unknown>;
  pivotCacheRecordsRaw: Record<string, unknown>;
  calcProperties: Record<string, unknown>;
}

/**
 * The top-level container for Excel workbook data.
 *
 * A Workbook manages worksheets, document metadata, images, and provides
 * IO helpers for reading/writing XLSX and CSV files.
 *
 * @example
 * ```ts
 * const workbook = new Workbook();
 * workbook.creator = 'ExcelTS';
 * workbook.created = new Date();
 *
 * const sheet = workbook.addWorksheet('Data');
 * sheet.getCell('A1').value = 'Hello';
 *
 * await workbook.xlsx.writeFile('output.xlsx');
 * ```
 */
class Workbook {
  /** Document category metadata */
  category: string;
  /** Company name metadata */
  company: string;
  /** Document creation date */
  created: Date;
  /** Document creator/author */
  creator?: string;
  /** Document description metadata */
  description: string;
  /** Document keywords metadata */
  keywords: string;
  /** Document manager metadata */
  manager: string;
  /** Last modification date */
  modified: Date;
  /** Last person to modify the document */
  lastModifiedBy?: string;
  /** Custom document properties */
  properties: Record<string, unknown>;
  /** Calculation properties for formulas */
  calcProperties: Record<string, unknown>;
  /** Document subject metadata */
  subject: string;
  /** Document title metadata */
  title: string;
  /** Workbook view settings */
  views: unknown[];
  /** Embedded media (images) */
  media: unknown[];
  /** @internal */
  pivotTables: unknown[];
  /** @internal */
  pivotTablesRaw: Record<string, unknown>;
  /** @internal */
  pivotCacheDefinitionsRaw: Record<string, unknown>;
  /** @internal */
  pivotCacheRecordsRaw: Record<string, unknown>;
  /** @internal */
  lastPrinted?: Date;
  /** @internal */
  language?: string;
  /** @internal */
  revision?: number;
  /** @internal */
  contentStatus?: string;
  /** @internal */
  private _worksheets: Worksheet[];
  /** @internal */
  private _definedNames: DefinedNames;
  /** @internal */
  private _xlsx?: XLSX;
  /** @internal */
  private _csv?: CSV;
  /** @internal */
  private _themes?: unknown;

  constructor() {
    this.category = '';
    this.company = '';
    this.created = new Date();
    this.description = '';
    this.keywords = '';
    this.manager = '';
    this.modified = this.created;
    this.properties = {};
    this.calcProperties = {};
    this._worksheets = [];
    this.subject = '';
    this.title = '';
    this.views = [];
    this.media = [];
    this.pivotTables = [];
    this.pivotTablesRaw = {};
    this.pivotCacheDefinitionsRaw = {};
    this.pivotCacheRecordsRaw = {};
    this._definedNames = new DefinedNames();
  }

  /**
   * XLSX file handler for reading and writing Excel files.
   *
   * @example
   * ```ts
   * // Read from file
   * await workbook.xlsx.readFile('input.xlsx');
   *
   * // Write to file
   * await workbook.xlsx.writeFile('output.xlsx');
   *
   * // Read from buffer
   * await workbook.xlsx.load(buffer);
   *
   * // Write to buffer
   * const buffer = await workbook.xlsx.writeBuffer();
   * ```
   */
  get xlsx(): XLSX {
    if (!this._xlsx) this._xlsx = new XLSX(this);
    return this._xlsx;
  }

  /**
   * CSV file handler for reading and writing CSV files.
   *
   * @example
   * ```ts
   * // Read from file
   * await workbook.csv.readFile('data.csv');
   *
   * // Write to file
   * await workbook.csv.writeFile('output.csv');
   * ```
   */
  get csv(): CSV {
    if (!this._csv) this._csv = new CSV(this);
    return this._csv;
  }

  /** @internal Next available worksheet ID */
  get nextId(): number {
    // find the next unique spot to add worksheet
    for (let i = 1; i < this._worksheets.length; i++) {
      if (!this._worksheets[i]) {
        return i;
      }
    }
    return this._worksheets.length || 1;
  }

  /**
   * Creates a new worksheet and adds it to the workbook.
   *
   * @param name - Worksheet name (max 31 chars, no special chars: * ? : \\ / [ ])
   * @param options - Optional worksheet configuration
   * @returns The newly created worksheet
   *
   * @example
   * ```ts
   * // Basic usage
   * const sheet = workbook.addWorksheet('Sales');
   *
   * // With tab color
   * const sheet = workbook.addWorksheet('Reports', {
   *   properties: { tabColor: { argb: 'FF00FF00' } }
   * });
   *
   * // With page setup
   * const sheet = workbook.addWorksheet('Print', {
   *   pageSetup: { orientation: 'landscape' }
   * });
   * ```
   */
  addWorksheet(name?: string, options?: WorksheetOptions): Worksheet {
    const id = this.nextId;

    // if options is a color, call it tabColor (and signal deprecated message)
    if (options) {
      if (typeof options === 'string') {
        // eslint-disable-next-line no-console
        console.trace(
          'tabColor argument is now deprecated. Please use workbook.addWorksheet(name, {properties: { tabColor: { argb: "rbg value" } }'
        );
        options = {
          properties: {
            tabColor: {argb: options},
          },
        };
      } else if (options.argb || options.theme || options.indexed) {
        // eslint-disable-next-line no-console
        console.trace(
          'tabColor argument is now deprecated. Please use workbook.addWorksheet(name, {properties: { tabColor: { ... } }'
        );
        options = {
          properties: {
            tabColor: options,
          },
        };
      }
    }

    const lastOrderNo = this._worksheets.reduce((acc, ws) => ((ws && ws.orderNo) > acc ? ws.orderNo : acc), 0);
    const worksheetOptions = Object.assign({}, options, {
      id,
      name,
      orderNo: lastOrderNo + 1,
      workbook: this,
    });

    const worksheet = new Worksheet(worksheetOptions);

    this._worksheets[id] = worksheet;
    return worksheet;
  }

  /** @internal */
  removeWorksheetEx(worksheet: Worksheet): void {
    delete this._worksheets[worksheet.id];
  }

  /**
   * Removes a worksheet from the workbook.
   *
   * @param id - Worksheet ID (number) or name (string)
   *
   * @example
   * ```ts
   * workbook.removeWorksheet(1);          // by ID
   * workbook.removeWorksheet('Sheet1');   // by name
   * ```
   */
  removeWorksheet(id: number | string): void {
    const worksheet = this.getWorksheet(id);
    if (worksheet) {
      worksheet.destroy();
    }
  }

  /**
   * Gets a worksheet by ID or name.
   *
   * @param id - Worksheet ID (number), name (string), or undefined for first sheet
   * @returns The worksheet, or undefined if not found
   *
   * @example
   * ```ts
   * const sheet = workbook.getWorksheet(1);        // by ID
   * const sheet = workbook.getWorksheet('Sales');  // by name
   * const first = workbook.getWorksheet();         // first sheet
   * ```
   */
  getWorksheet(id?: number | string): Worksheet | undefined {
    if (id === undefined) {
      return this._worksheets.find(Boolean);
    }
    if (typeof id === 'number') {
      return this._worksheets[id];
    }
    if (typeof id === 'string') {
      return this._worksheets.find(worksheet => worksheet && worksheet.name === id);
    }
    return undefined;
  }

  /**
   * All worksheets in the workbook, sorted by display order.
   *
   * @example
   * ```ts
   * for (const sheet of workbook.worksheets) {
   *   console.log(sheet.name);
   * }
   * ```
   */
  get worksheets(): Worksheet[] {
    // return a clone of _worksheets
    return this._worksheets
      .slice(1)
      .sort((a, b) => a.orderNo - b.orderNo)
      .filter(Boolean);
  }

  /**
   * Iterates over all worksheets in the workbook.
   *
   * @param iteratee - Callback function receiving (worksheet, id)
   *
   * @example
   * ```ts
   * workbook.eachSheet((sheet, id) => {
   *   console.log(`Sheet ${id}: ${sheet.name}`);
   * });
   * ```
   */
  eachSheet(iteratee: (worksheet: Worksheet, id: number) => void): void {
    this.worksheets.forEach(sheet => {
      iteratee(sheet, sheet.id);
    });
  }

  /**
   * Defined names (named ranges) in the workbook.
   *
   * @example
   * ```ts
   * // Add a named range
   * workbook.definedNames.add('MyRange', 'Sheet1!$A$1:$B$10');
   * ```
   */
  get definedNames(): DefinedNames {
    return this._definedNames;
  }

  /** @internal Clears theme data */
  clearThemes(): void {
    // Note: themes are not an exposed feature, meddle at your peril!
    this._themes = undefined;
  }

  /**
   * Adds an image to the workbook for later use in worksheets.
   *
   * @param image - Image data with filename, extension, and buffer
   * @returns Image ID to reference when adding to worksheets
   *
   * @example
   * ```ts
   * import fs from 'fs';
   *
   * const imageId = workbook.addImage({
   *   buffer: fs.readFileSync('logo.png'),
   *   extension: 'png',
   * });
   *
   * // Then add to worksheet
   * worksheet.addImage(imageId, 'A1:C5');
   * ```
   */
  addImage(image: WorkbookImage): number {
    // TODO:  validation?
    const id = this.media.length;
    this.media.push(Object.assign({}, image, {type: 'image'}));
    return id;
  }

  /**
   * Gets an image by its ID.
   *
   * @param id - Image ID returned from addImage()
   * @returns The image data
   */
  getImage(id: number): unknown {
    return this.media[id];
  }

  get model(): WorkbookModel {
    return {
      creator: this.creator || 'Unknown',
      lastModifiedBy: this.lastModifiedBy || 'Unknown',
      lastPrinted: this.lastPrinted,
      created: this.created,
      modified: this.modified,
      properties: this.properties,
      worksheets: this.worksheets.map(worksheet => worksheet.model),
      sheets: this.worksheets.map(ws => ws.model).filter(Boolean),
      definedNames: this._definedNames.model,
      views: this.views,
      company: this.company,
      manager: this.manager,
      title: this.title,
      subject: this.subject,
      keywords: this.keywords,
      category: this.category,
      description: this.description,
      language: this.language,
      revision: this.revision,
      contentStatus: this.contentStatus,
      themes: this._themes,
      media: this.media,
      pivotTables: this.pivotTables,
      pivotTablesRaw: this.pivotTablesRaw,
      pivotCacheDefinitionsRaw: this.pivotCacheDefinitionsRaw,
      pivotCacheRecordsRaw: this.pivotCacheRecordsRaw,
      calcProperties: this.calcProperties,
    };
  }

  set model(value: WorkbookModel) {
    this.creator = value.creator;
    this.lastModifiedBy = value.lastModifiedBy;
    this.lastPrinted = value.lastPrinted;
    this.created = value.created;
    this.modified = value.modified;
    this.company = value.company;
    this.manager = value.manager;
    this.title = value.title;
    this.subject = value.subject;
    this.keywords = value.keywords;
    this.category = value.category;
    this.description = value.description;
    this.language = value.language;
    this.revision = value.revision;
    this.contentStatus = value.contentStatus;

    this.properties = value.properties;
    this.calcProperties = value.calcProperties;
    this._worksheets = [];
    value.worksheets.forEach(worksheetModel => {
      const {id, name, state} = worksheetModel;
      const orderNo = value.sheets && value.sheets.findIndex(ws => ws.id === id);
      const worksheet = (this._worksheets[id] = new Worksheet({
        id,
        name,
        orderNo,
        state,
        workbook: this,
      }));
      worksheet.model = worksheetModel;
    });

    this._definedNames.model = value.definedNames;
    this.views = value.views;
    this._themes = value.themes;
    this.media = value.media || [];
    this.pivotTables = value.pivotTables || [];
    this.pivotTablesRaw = value.pivotTablesRaw || {};
    this.pivotCacheDefinitionsRaw = value.pivotCacheDefinitionsRaw || {};
    this.pivotCacheRecordsRaw = value.pivotCacheRecordsRaw || {};
  }
}

export default Workbook;