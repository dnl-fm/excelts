
import _ from '../utils/under-dash.ts';
import colCache from '../utils/col-cache.ts';
import Range from './range.ts';
import Row from './row.ts';
import Column from './column.ts';
import Enums from './enums.ts';
import Image from './image.ts';
import Table from './table.ts';
import DataValidations from './data-validations.ts';
import Encryptor from '../utils/encryptor.ts';
import { toString as bytesToString } from '../utils/bytes.ts';
import { makePivotTable } from './pivot-table.ts';
import copyStyle from '../utils/copy-style.ts';
import type Workbook from './workbook.ts';
import type Cell from './cell.ts';

/** Worksheet constructor options */
export interface WorksheetOptions {
  workbook?: Workbook;
  id?: number;
  orderNo?: number;
  name?: string;
  state?: 'visible' | 'hidden' | 'veryHidden';
  properties?: Record<string, unknown>;
  pageSetup?: Record<string, unknown>;
  headerFooter?: Record<string, unknown>;
  views?: unknown[];
  autoFilter?: unknown;
}

/** Column definition for worksheet */
export interface ColumnDefinition {
  header?: string | string[];
  key?: string;
  width?: number;
  style?: Record<string, unknown>;
  hidden?: boolean;
  outlineLevel?: number;
}

/** Worksheet model for serialization */
export interface WorksheetModel {
  id: number;
  name: string;
  dataValidations: unknown;
  properties: Record<string, unknown>;
  state: 'visible' | 'hidden' | 'veryHidden';
  pageSetup: Record<string, unknown>;
  headerFooter: Record<string, unknown>;
  rowBreaks: unknown[];
  views: unknown[];
  autoFilter: unknown;
  media: unknown[];
  sheetProtection: unknown;
  tables: unknown[];
  pivotTables: unknown[];
  conditionalFormattings: unknown[];
  cols?: unknown;
  rows?: unknown[];
  dimensions?: Range;
  merges?: string[];
  mergeCells?: string[];
}

/**
 * A worksheet within a workbook, containing rows, columns, and cells.
 *
 * Provides access to cell data, formatting, merged cells, images, tables,
 * data validation, and page layout settings.
 *
 * @example
 * ```ts
 * const sheet = workbook.addWorksheet('Data');
 *
 * // Set column headers
 * sheet.columns = [
 *   { header: 'Name', key: 'name', width: 20 },
 *   { header: 'Age', key: 'age', width: 10 },
 * ];
 *
 * // Add data rows
 * sheet.addRow({ name: 'Alice', age: 30 });
 * sheet.addRow({ name: 'Bob', age: 25 });
 *
 * // Access cells
 * sheet.getCell('A1').font = { bold: true };
 * ```
 */
class Worksheet {
  id: number;
  orderNo: number;
  state: 'visible' | 'hidden' | 'veryHidden';
  properties: Record<string, unknown>;
  pageSetup: Record<string, unknown>;
  headerFooter: Record<string, unknown>;
  dataValidations: DataValidations;
  views: unknown[];
  autoFilter: unknown;
  sheetProtection: unknown;
  tables: Record<string, Table>;
  pivotTables: unknown[];
  conditionalFormattings: unknown[];
  rowBreaks: unknown[];
  private _workbook: Workbook;
  private _name: string;
  private _rows: Row[];
  private _columns: Column[] | null;
  private _keys: Record<string, Column>;
  private _merges: Record<string, Range>;
  private _media: Image[];
  private _headerRowCount?: number;

  constructor(options?: WorksheetOptions) {
    options = options || {};
    this._workbook = options.workbook;

    // in a workbook, each sheet will have a number
    this.id = options.id;
    this.orderNo = options.orderNo;

    // and a name
    this.name = options.name;

    // add a state
    this.state = options.state || 'visible';

    // rows allows access organised by row. Sparse array of arrays indexed by row-1, col
    // Note: _rows is zero based. Must subtract 1 to go from cell.row to index
    this._rows = [];

    // column definitions
    this._columns = null;

    // column keys (addRow convenience): key ==> this._collumns index
    this._keys = {};

    // keep record of all merges
    this._merges = {};

    // record of all row and column pageBreaks
    this.rowBreaks = [];

    // for tabColor, default row height, outline levels, etc
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

    // for all things printing
    this.pageSetup = Object.assign(
      {},
      {
        margins: {left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3},
        orientation: 'portrait',
        horizontalDpi: 4294967295,
        verticalDpi: 4294967295,
        fitToPage: !!(
          options.pageSetup &&
          (options.pageSetup.fitToWidth || options.pageSetup.fitToHeight) &&
          !options.pageSetup.scale
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
        firstPageNumber: undefined,
        horizontalCentered: false,
        verticalCentered: false,
        rowBreaks: null,
        colBreaks: null,
      },
      options.pageSetup
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

    this.dataValidations = new DataValidations();

    // for freezepanes, split, zoom, gridlines, etc
    this.views = options.views || [];

    this.autoFilter = options.autoFilter || null;

    // for images, etc
    this._media = [];

    // worksheet protection
    this.sheetProtection = null;

    // for tables
    this.tables = {};

    this.pivotTables = [];

    this.conditionalFormattings = [];
  }

  get name(): string {
    return this._name;
  }

  set name(name: string | undefined) {
    if (name === undefined) {
      name = `sheet${this.id}`;
    }

    if (this._name === name) return;

    if (typeof name !== 'string') {
      throw new Error('The name has to be a string.');
    }

    if (name === '') {
      throw new Error('The name can\'t be empty.');
    }

    if (name === 'History') {
      throw new Error('The name "History" is protected. Please use a different name.');
    }

    // Illegal character in worksheet name: asterisk (*), question mark (?),
    // colon (:), forward slash (/ \), or bracket ([])
    if (/[*?:/\\[\]]/.test(name)) {
      throw new Error(`Worksheet name ${name} cannot include any of the following characters: * ? : \\ / [ ]`);
    }

    if (/(^')|('$)/.test(name)) {
      throw new Error(`The first or last character of worksheet name cannot be a single quotation mark: ${name}`);
    }

    if (name && name.length > 31) {
      // eslint-disable-next-line no-console
      console.warn(`Worksheet name ${name} exceeds 31 chars. This will be truncated`);
      name = name.substring(0, 31);
    }

    if (this._workbook.worksheets.find(ws => ws && ws.name.toLowerCase() === name.toLowerCase())) {
      throw new Error(`Worksheet name already exists: ${name}`);
    }

    this._name = name;
  }

  /** The parent workbook containing this worksheet */
  get workbook(): Workbook {
    return this._workbook;
  }

  /**
   * Removes this worksheet from its parent workbook.
   * Call when you're done with this worksheet.
   */
  destroy(): void {
    this._workbook.removeWorksheetEx(this);
  }

  /**
   * The bounding range of all cells with data in this worksheet.
   *
   * @example
   * ```ts
   * const dims = sheet.dimensions;
   * console.log(`Data from ${dims.top},${dims.left} to ${dims.bottom},${dims.right}`);
   * ```
   */
  get dimensions(): Range {
    const dimensions = new Range();
    this._rows.forEach(row => {
      if (row) {
        const rowDims = row.dimensions;
        if (rowDims) {
          dimensions.expand(row.number, rowDims.min, row.number, rowDims.max);
        }
      }
    });
    return dimensions;
  }

  // =========================================================================
  // Columns

  /**
   * Column definitions for this worksheet.
   *
   * @example
   * ```ts
   * sheet.columns = [
   *   { header: 'ID', key: 'id', width: 10 },
   *   { header: 'Name', key: 'name', width: 30 },
   *   { header: 'Email', key: 'email', width: 40 },
   * ];
   * ```
   */
  get columns(): Column[] | null {
    return this._columns;
  }

  /**
   * Sets columns from an array of column definitions.
   * Note: any headers defined will overwrite existing cell values.
   */
  set columns(value: ColumnDefinition[]) {
    // calculate max header row count
    this._headerRowCount = value.reduce((pv, cv) => {
      let headerCount = 0;
      if (cv.header) {
        headerCount = typeof cv.header === 'string' ? 1 : cv.header.length;
      }
      return Math.max(pv, headerCount);
    }, 0);

    // First pass: create Column objects without setting defn
    // This prevents header setter from creating duplicate columns via getColumn()
    let count = 1;
    const columns = (this._columns = []);
    value.forEach(() => {
      const column = new Column(this, count++, false);
      columns.push(column);
    });

    // Second pass: set definitions (which sets headers, keys, etc.)
    value.forEach((defn, index) => {
      columns[index].defn = defn as Record<string, unknown>;
    });
  }

  getColumnKey(key: string): Column | undefined {
    return this._keys[key];
  }

  setColumnKey(key: string, value: Column): void {
    this._keys[key] = value;
  }

  deleteColumnKey(key: string): void {
    delete this._keys[key];
  }

  eachColumnKey(f: (column: Column, key: string) => void): void {
    _.each(this._keys, f);
  }

  /**
   * Gets a column by number, letter, or key. Creates the column if it doesn't exist.
   *
   * @param c - Column number (1-based), letter ('A', 'B', etc.), or key name
   * @returns The column object
   *
   * @example
   * ```ts
   * const col = sheet.getColumn(1);      // by number
   * const col = sheet.getColumn('A');    // by letter
   * const col = sheet.getColumn('name'); // by key
   *
   * col.width = 20;
   * col.hidden = true;
   * ```
   */
  getColumn(c: number | string): Column {
    let colNum: number;
    if (typeof c === 'string') {
      // if it matches a key'd column, return that
      const col = this._keys[c];
      if (col) return col;

      // otherwise, assume letter
      colNum = colCache.l2n(c);
    } else {
      colNum = c;
    }
    if (!this._columns) {
      this._columns = [];
    }
    if (colNum > this._columns.length) {
      let n = this._columns.length + 1;
      while (n <= colNum) {
        this._columns.push(new Column(this, n++));
      }
    }
    return this._columns[colNum - 1];
  }

  spliceColumns(start: number, count: number, ...inserts: unknown[][]): void {
    const rows = this._rows;
    const nRows = rows.length;
    if (inserts.length > 0) {
      // must iterate over all rows whether they exist yet or not
      for (let i = 0; i < nRows; i++) {
        const rowArguments: unknown[] = [start, count];
        // eslint-disable-next-line no-loop-func
        inserts.forEach(insert => {
          rowArguments.push(insert[i] || null);
        });
        const row = this.getRow(i + 1);
        // eslint-disable-next-line prefer-spread
        row.splice.apply(row, rowArguments as [number, number, ...unknown[]]);
      }
    } else {
      // nothing to insert, so just splice all rows
      this._rows.forEach(r => {
        if (r) {
          r.splice(start, count);
        }
      });
    }

    // splice column definitions
    const nExpand = inserts.length - count;
    const nKeep = start + count;
    const nEnd = this._columns?.length || 0;
    if (nExpand < 0) {
      for (let i = start + inserts.length; i <= nEnd; i++) {
        this.getColumn(i).defn = this.getColumn(i - nExpand).defn;
      }
    } else if (nExpand > 0) {
      for (let i = nEnd; i >= nKeep; i--) {
        this.getColumn(i + nExpand).defn = this.getColumn(i).defn;
      }
    }
    for (let i = start; i < start + inserts.length; i++) {
      this.getColumn(i).defn = null;
    }

    // account for defined names
    this.workbook.definedNames.spliceColumns(this.name, start, count, inserts.length);
  }

  get lastColumn(): Column {
    return this.getColumn(this.columnCount);
  }

  get columnCount(): number {
    let maxCount = 0;
    this.eachRow(row => {
      maxCount = Math.max(maxCount, row.cellCount);
    });
    return maxCount;
  }

  get actualColumnCount(): number {
    // performance nightmare - for each row, counts all the columns used
    const counts: boolean[] = [];
    let count = 0;
    this.eachRow(row => {
      row.eachCell(({col}: { col: number }) => {
        if (!counts[col]) {
          counts[col] = true;
          count++;
        }
      });
    });
    return count;
  }

  // =========================================================================
  // Rows

  _commitRow(_row?: unknown): void {
    // nop - allows streaming reader to fill a document
  }

  get _lastRowNumber(): number {
    // need to cope with results of splice
    const rows = this._rows;
    let n = rows.length;
    while (n > 0 && rows[n - 1] === undefined) {
      n--;
    }
    return n;
  }

  get _nextRow(): number {
    return this._lastRowNumber + 1;
  }

  get lastRow(): Row | undefined {
    if (this._rows.length) {
      return this._rows[this._rows.length - 1];
    }
    return undefined;
  }

  /**
   * Finds an existing row by number. Returns undefined if the row doesn't exist.
   *
   * @param r - Row number (1-based)
   * @returns The row, or undefined
   */
  findRow(r: number): Row | undefined {
    return this._rows[r - 1];
  }

  /**
   * Finds multiple existing rows starting at a position.
   *
   * @param start - Starting row number (1-based)
   * @param length - Number of rows to find
   * @returns Array of rows (may contain undefined for non-existent rows)
   */
  findRows(start: number, length: number): Row[] {
    return this._rows.slice(start - 1, start - 1 + length);
  }

  /** Total number of rows in the worksheet (including empty rows) */
  get rowCount(): number {
    return this._lastRowNumber;
  }

  /** Number of rows that contain actual data */
  get actualRowCount(): number {
    // counts actual rows that have actual data
    let count = 0;
    this.eachRow(() => {
      count++;
    });
    return count;
  }

  /**
   * Gets a row by number. Creates the row if it doesn't exist.
   *
   * @param r - Row number (1-based)
   * @returns The row object
   *
   * @example
   * ```ts
   * const row = sheet.getRow(1);
   * row.getCell(1).value = 'Hello';
   * row.height = 20;
   * ```
   */
  getRow(r: number): Row {
    let row = this._rows[r - 1];
    if (!row) {
      row = this._rows[r - 1] = new Row(this, r);
    }
    return row;
  }

  /**
   * Gets multiple consecutive rows. Creates rows that don't exist.
   *
   * @param start - Starting row number (1-based)
   * @param length - Number of rows to get
   * @returns Array of row objects
   */
  getRows(start: number, length: number): Row[] | undefined {
    if (length < 1) return undefined;
    const rows: Row[] = [];
    for (let i = start; i < start + length; i++) {
      rows.push(this.getRow(i));
    }
    return rows;
  }

  /**
   * Adds a new row at the end of the worksheet.
   *
   * @param value - Row data as array or object with column keys
   * @param style - Style option: 'n' (none), 'i' (inherit from above), 'o' (inherit from below)
   * @returns The newly created row
   *
   * @example
   * ```ts
   * // Array values (columns A, B, C, ...)
   * sheet.addRow([1, 'Alice', 'alice@example.com']);
   *
   * // Object values (requires column keys)
   * sheet.addRow({ id: 1, name: 'Alice', email: 'alice@example.com' });
   * ```
   */
  addRow(value: unknown[] | Record<string, unknown>, style: string = 'n'): Row {
    const rowNo = this._nextRow;
    const row = this.getRow(rowNo);
    row.values = value;
    this._setStyleOption(rowNo, style[0] === 'i' ? style : 'n');
    return row;
  }

  /**
   * Adds multiple rows at the end of the worksheet.
   *
   * @param value - Array of row data
   * @param style - Style option: 'n' (none), 'i' (inherit from above), 'o' (inherit from below)
   * @returns Array of newly created rows
   *
   * @example
   * ```ts
   * sheet.addRows([
   *   { id: 1, name: 'Alice' },
   *   { id: 2, name: 'Bob' },
   *   { id: 3, name: 'Charlie' },
   * ]);
   * ```
   */
  addRows(value: Array<unknown[] | Record<string, unknown>>, style: string = 'n'): Row[] {
    const rows: Row[] = [];
    value.forEach(row => {
      rows.push(this.addRow(row, style));
    });
    return rows;
  }

  /**
   * Inserts a row at a specific position, shifting existing rows down.
   *
   * @param pos - Row position to insert at (1-based)
   * @param value - Row data
   * @param style - Style option
   * @returns The inserted row
   */
  insertRow(pos: number, value: unknown[] | Record<string, unknown>, style: string = 'n'): Row {
    this.spliceRows(pos, 0, value);
    this._setStyleOption(pos, style);
    return this.getRow(pos);
  }

  /**
   * Inserts multiple rows at a specific position, shifting existing rows down.
   *
   * @param pos - Row position to insert at (1-based)
   * @param values - Array of row data
   * @param style - Style option
   * @returns Array of inserted rows
   */
  insertRows(pos: number, values: Array<unknown[] | Record<string, unknown>>, style: string = 'n'): Row[] | undefined {
    this.spliceRows(pos, 0, ...values);
    if (style !== 'n') {
      // copy over the styles
      for (let i = 0; i < values.length; i++) {
        if (style[0] === 'o' && this.findRow(values.length + pos + i) !== undefined) {
          this._copyStyle(values.length + pos + i, pos + i, style[1] === '+');
        } else if (style[0] === 'i' && this.findRow(pos - 1) !== undefined) {
          this._copyStyle(pos - 1, pos + i, style[1] === '+');
        }
      }
    }
    return this.getRows(pos, values.length);
  }

  // set row at position to same style as of either pervious row (option 'i') or next row (option 'o')
  private _setStyleOption(pos: number, style: string = 'n'): void {
    if (style[0] === 'o' && this.findRow(pos + 1) !== undefined) {
      this._copyStyle(pos + 1, pos, style[1] === '+');
    } else if (style[0] === 'i' && this.findRow(pos - 1) !== undefined) {
      this._copyStyle(pos - 1, pos, style[1] === '+');
    }
  }

  private _copyStyle(src: number, dest: number, styleEmpty: boolean = false): void {
    const rSrc = this.getRow(src);
    const rDst = this.getRow(dest);
    rDst.style = copyStyle(rSrc.style);
    // eslint-disable-next-line no-loop-func
    rSrc.eachCell({includeEmpty: styleEmpty}, (cell: Cell, colNumber: number) => {
      rDst.getCell(colNumber).style = copyStyle(cell.style);
    });
    rDst.height = rSrc.height;
  }

  duplicateRow(rowNum: number, count: number, insert: boolean = false): void {
    // create count duplicates of rowNum
    // either inserting new or overwriting existing rows

    const rSrc = this._rows[rowNum - 1];
    const inserts = Array.from({ length: count }, () => rSrc.values);
    this.spliceRows(rowNum + 1, insert ? 0 : count, ...inserts);

    // now copy styles...
    for (let i = 0; i < count; i++) {
      const rDst = this._rows[rowNum + i];
      rDst.style = rSrc.style;
      rDst.height = rSrc.height;
      // eslint-disable-next-line no-loop-func
      rSrc.eachCell({includeEmpty: true}, (cell: Cell, colNumber: number) => {
        rDst.getCell(colNumber).style = cell.style;
      });
    }
  }

  spliceRows(start: number, count: number, ...inserts: unknown[]): void {
    // same problem as row.splice, except worse.
    const nKeep = start + count;
    const nInserts = inserts.length;
    const nExpand = nInserts - count;
    const nEnd = this._rows.length;
    let i;
    let rSrc;
    if (nExpand < 0) {
      // remove rows
      if (start === nEnd) {
        this._rows[nEnd - 1] = undefined;
      }
      for (i = nKeep; i <= nEnd; i++) {
        rSrc = this._rows[i - 1];
        if (rSrc) {
          const rDst = this.getRow(i + nExpand);
          rDst.values = rSrc.values;
          rDst.style = rSrc.style;
          rDst.height = rSrc.height;
          // eslint-disable-next-line no-loop-func
          rSrc.eachCell({includeEmpty: true}, (cell, colNumber) => {
            rDst.getCell(colNumber).style = cell.style;
          });
          this._rows[i - 1] = undefined;
        } else {
          this._rows[i + nExpand - 1] = undefined;
        }
      }
    } else if (nExpand > 0) {
      // insert new cells
      for (i = nEnd; i >= nKeep; i--) {
        rSrc = this._rows[i - 1];
        if (rSrc) {
          const rDst = this.getRow(i + nExpand);
          rDst.values = rSrc.values;
          rDst.style = rSrc.style;
          rDst.height = rSrc.height;
          // eslint-disable-next-line no-loop-func
          rSrc.eachCell({includeEmpty: true}, (cell, colNumber) => {
            rDst.getCell(colNumber).style = cell.style;

            // remerge cells accounting for insert offset
            if (cell._value.constructor.name === 'MergeValue') {
              const cellToBeMerged = this.getRow(cell._row._number + nInserts).getCell(colNumber);
              const prevMaster = cell._value._master;
              const newMaster = this.getRow(prevMaster._row._number + nInserts).getCell(prevMaster._column._number);
              cellToBeMerged.merge(newMaster);
            }
          });
        } else {
          this._rows[i + nExpand - 1] = undefined;
        }
      }
    }

    // now copy over the new values
    for (i = 0; i < nInserts; i++) {
      const rDst = this.getRow(start + i);
      rDst.style = {};
      rDst.values = inserts[i] as unknown[] | Record<string, unknown> | undefined;
    }

    // account for defined names
    this.workbook.definedNames.spliceRows(this.name, start, count, nInserts);
  }

  /**
   * Iterates over rows in the worksheet.
   *
   * @param options - Options object with `includeEmpty: boolean`, or the iteratee function
   * @param iteratee - Callback function receiving (row, rowNumber)
   *
   * @example
   * ```ts
   * // Only rows with data
   * sheet.eachRow((row, rowNumber) => {
   *   console.log(`Row ${rowNumber}: ${row.values}`);
   * });
   *
   * // Include empty rows
   * sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
   *   console.log(`Row ${rowNumber}`);
   * });
   * ```
   */
  eachRow(options: unknown, iteratee?: unknown): void {
    if (!iteratee) {
      iteratee = options;
      options = undefined;
    }
    if (options && options.includeEmpty) {
      const n = this._rows.length;
      for (let i = 1; i <= n; i++) {
        iteratee(this.getRow(i), i);
      }
    } else {
      this._rows.forEach(row => {
        if (row && row.hasValues) {
          iteratee(row, row.number);
        }
      });
    }
  }

  /**
   * Returns all row values as a sparse array.
   *
   * @returns Sparse array where index is row number and value is row data
   */
  getSheetValues(): unknown[] {
    const rows: unknown[] = [];
    this._rows.forEach(row => {
      if (row) {
        rows[row.number] = row.values;
      }
    });
    return rows;
  }

  // =========================================================================
  // Cells

  /**
   * Finds an existing cell. Returns undefined if the cell doesn't exist.
   *
   * @param r - Row number or cell address (e.g., 'A1')
   * @param c - Column number (if r is row number)
   * @returns The cell, or undefined
   */
  findCell(r: number | string, c?: number): Cell | undefined {
    const address = colCache.getAddress(r, c);
    const row = this._rows[address.row - 1];
    return row ? row.findCell(address.col) : undefined;
  }

  /**
   * Gets a cell by address. Creates the cell if it doesn't exist.
   *
   * @param r - Row number or cell address (e.g., 'A1', 'B2')
   * @param c - Column number (if r is row number)
   * @returns The cell object
   *
   * @example
   * ```ts
   * // By address
   * sheet.getCell('A1').value = 'Hello';
   * sheet.getCell('B2').value = 42;
   *
   * // By row and column
   * sheet.getCell(1, 1).value = 'Hello';
   * sheet.getCell(2, 2).value = 42;
   * ```
   */
  getCell(r: number | string, c?: number): Cell {
    const address = colCache.getAddress(r, c);
    const row = this.getRow(address.row);
    return row.getCellEx(address);
  }

  // =========================================================================
  // Merge

  /**
   * Merges a range of cells into a single cell.
   *
   * @param cells - Range as 'A1:B2', ['A1', 'B2'], or [top, left, bottom, right]
   *
   * @example
   * ```ts
   * sheet.mergeCells('A1:B2');           // by range string
   * sheet.mergeCells('A1', 'B2');        // by corners
   * sheet.mergeCells(1, 1, 2, 2);        // by coordinates
   * ```
   */
  mergeCells(...cells: unknown[]): void {
    const dimensions = new Range(cells);
    this._mergeCellsInternal(dimensions);
  }

  /**
   * Merges cells without copying styles from the master cell.
   *
   * @param cells - Range specification (same as mergeCells)
   */
  mergeCellsWithoutStyle(...cells: unknown[]): void {
    const dimensions = new Range(cells);
    this._mergeCellsInternal(dimensions, true);
  }

  _mergeCellsInternal(dimensions: Range, ignoreStyle?: boolean): void {
    // check cells aren't already merged
    _.each(this._merges, merge => {
      if (merge.intersects(dimensions)) {
        throw new Error('Cannot merge already merged cells');
      }
    });

    // apply merge
    const master = this.getCell(dimensions.top, dimensions.left);
    for (let i = dimensions.top; i <= dimensions.bottom; i++) {
      for (let j = dimensions.left; j <= dimensions.right; j++) {
        // merge all but the master cell
        if (i > dimensions.top || j > dimensions.left) {
          this.getCell(i, j).merge(master, ignoreStyle);
        }
      }
    }

    // index merge
    this._merges[master.address] = dimensions;
  }

  _unMergeMaster(master: { address: string }): void {
    // master is always top left of a rectangle
    const merge = this._merges[master.address];
    if (merge) {
      for (let i = merge.top; i <= merge.bottom; i++) {
        for (let j = merge.left; j <= merge.right; j++) {
          this.getCell(i, j).unmerge();
        }
      }
      delete this._merges[master.address];
    }
  }

  /** Whether this worksheet contains any merged cells */
  get hasMerges(): boolean {
    // return true if this._merges has a merge object
    return _.some(this._merges, Boolean);
  }

  /**
   * Unmerges cells in a range. If any cell in the range is part of a merge,
   * the entire merge is removed.
   *
   * @param cells - Range specification (same as mergeCells)
   *
   * @example
   * ```ts
   * sheet.unMergeCells('A1:B2');
   * ```
   */
  unMergeCells(...cells: unknown[]): void {
    const dimensions = new Range(cells);

    // find any cells in that range and unmerge them
    for (let i = dimensions.top; i <= dimensions.bottom; i++) {
      for (let j = dimensions.left; j <= dimensions.right; j++) {
        const cell = this.findCell(i, j);
        if (cell) {
          if (cell.type === Enums.ValueType.Merge) {
            // this cell merges to another master
            this._unMergeMaster(cell.master);
          } else if (this._merges[cell.address]) {
            // this cell is a master
            this._unMergeMaster(cell);
          }
        }
      }
    }
  }

  // ===========================================================================
  // Shared/Array Formula
  fillFormula(range: string, formula: string, results?: unknown, shareType = 'shared') {
    // Define formula for top-left cell and share to rest
    const decoded = colCache.decode(range) as { top: number; left: number; bottom: number; right: number };
    const {top, left, bottom, right} = decoded;
    const width = right - left + 1;
    const masterAddress = colCache.encodeAddress(top, left);
    const isShared = shareType === 'shared';

    // work out result accessor
    let getResult;
    if (typeof results === 'function') {
      getResult = results;
    } else if (Array.isArray(results)) {
      if (Array.isArray(results[0])) {
        getResult = (row, col) => results[row - top][col - left];
      } else {
        // eslint-disable-next-line no-mixed-operators
        getResult = (row, col) => results[(row - top) * width + (col - left)];
      }
    } else {
      getResult = () => undefined;
    }
    let first = true;
    for (let r = top; r <= bottom; r++) {
      for (let c = left; c <= right; c++) {
        if (first) {
          this.getCell(r, c).value = {
            shareType,
            formula,
            ref: range,
            result: getResult(r, c),
          };
          first = false;
        } else {
          this.getCell(r, c).value = isShared
            ? {
                sharedFormula: masterAddress,
                result: getResult(r, c),
              }
            : getResult(r, c);
        }
      }
    }
  }

  // =========================================================================
  // Images

  /**
   * Adds an image to the worksheet.
   *
   * @param imageId - Image ID from workbook.addImage()
   * @param range - Position as range string 'A1:C5' or anchor object
   *
   * @example
   * ```ts
   * const imageId = workbook.addImage({
   *   buffer: imageBuffer,
   *   extension: 'png',
   * });
   *
   * // Position by cell range
   * sheet.addImage(imageId, 'A1:C5');
   *
   * // Position with detailed anchoring
   * sheet.addImage(imageId, {
   *   tl: { col: 0, row: 0 },
   *   br: { col: 3, row: 5 },
   * });
   * ```
   */
  addImage(imageId: unknown, range: unknown): void {
    const model = {
      type: 'image',
      imageId,
      range,
    };
    this._media.push(new Image(this, model));
  }

  /**
   * Gets all images in this worksheet.
   *
   * @returns Array of image objects
   */
  getImages(): Image[] {
    return this._media.filter(m => m.type === 'image') as Image[];
  }

  /**
   * Sets the worksheet background image.
   *
   * @param imageId - Image ID from workbook.addImage()
   */
  addBackgroundImage(imageId: string | number): void {
    const model = {
      type: 'background',
      imageId,
    };
    this._media.push(new Image(this, model));
  }

  /**
   * Gets the background image ID, if set.
   *
   * @returns Image ID or undefined
   */
  getBackgroundImageId(): string | number | undefined {
    const image = this._media.find(m => m.type === 'background');
    return image ? (image.imageId as string | number) : undefined;
  }

  // =========================================================================
  // Worksheet Protection

  /**
   * Protects the worksheet with an optional password.
   *
   * @param password - Optional password to protect the sheet
   * @param options - Protection options (selectLockedCells, selectUnlockedCells, etc.)
   * @returns Promise that resolves when protection is applied
   *
   * @example
   * ```ts
   * // Basic protection
   * await sheet.protect('password123');
   *
   * // With options
   * await sheet.protect('password123', {
   *   selectLockedCells: false,
   *   formatColumns: true,
   * });
   * ```
   */
  async protect(password?: string, options?: Record<string, unknown>): Promise<void> {
    // TODO: make this function truly async
    // perhaps marshal to worker thread or something
    const protection: Record<string, unknown> = {
      sheet: true,
    };
    if (options && 'spinCount' in options) {
      // force spinCount to be integer >= 0
      options.spinCount = Number.isFinite(options.spinCount as number) ? Math.round(Math.max(0, options.spinCount as number)) : 100000;
    }
    if (password) {
      protection.algorithmName = 'SHA-512';
      protection.saltValue = bytesToString(Encryptor.randomBytes(16), 'base64');
      protection.spinCount = options && 'spinCount' in options ? (options.spinCount as number) : 100000; // allow user specified spinCount
      protection.hashValue = await Encryptor.convertPasswordToHash(
        password,
        'SHA512',
        protection.saltValue as string,
        protection.spinCount as number
      );
    }
    if (options) {
      this.sheetProtection = Object.assign(protection, options);
      if (!password && 'spinCount' in options) {
        delete (this.sheetProtection as Record<string, unknown>).spinCount;
      }
    } else {
      this.sheetProtection = protection;
    }
  }

  /**
   * Removes worksheet protection.
   */
  unprotect() {
    this.sheetProtection = null;
  }

  // =========================================================================
  // Tables

  /**
   * Adds a table to the worksheet.
   *
   * @param model - Table definition with name, ref, columns, and rows
   * @returns The created table
   *
   * @example
   * ```ts
   * sheet.addTable({
   *   name: 'SalesTable',
   *   ref: 'A1',
   *   headerRow: true,
   *   columns: [
   *     { name: 'Product', filterButton: true },
   *     { name: 'Quantity', filterButton: true },
   *     { name: 'Price', filterButton: true },
   *   ],
   *   rows: [
   *     ['Widgets', 100, 9.99],
   *     ['Gadgets', 50, 19.99],
   *   ],
   * });
   * ```
   */
  addTable(model: Record<string, unknown>): Table {
    const table = new Table(this, model as any);
    this.tables[model.name as string] = table;
    return table;
  }

  /**
   * Gets a table by name.
   *
   * @param name - Table name
   * @returns The table, or undefined
   */
  getTable(name: string): Table | undefined {
    return this.tables[name];
  }

  /**
   * Removes a table by name.
   *
   * @param name - Table name
   */
  removeTable(name: string): void {
    delete this.tables[name];
  }

  /**
   * Gets all tables in this worksheet.
   *
   * @returns Array of table objects
   */
  getTables(): Table[] {
    return Object.values(this.tables);
  }

  // =========================================================================
  // Pivot Tables
  addPivotTable(model: Record<string, unknown>): unknown {
    // eslint-disable-next-line no-console
    console.warn(
      `Warning: Pivot Table support is experimental. 
Please leave feedback at https://github.com/exceljs/exceljs/discussions/2575`
    );

    const pivotTable = makePivotTable(this, model);

    this.pivotTables.push(pivotTable);
    (this.workbook as any).pivotTables = [
      ...((this.workbook as any).pivotTables as unknown[] || []),
      pivotTable
    ];

    return pivotTable;
  }

  // ===========================================================================
  // Conditional Formatting

  /**
   * Adds a conditional formatting rule to the worksheet.
   *
   * @param cf - Conditional formatting rule
   *
   * @example
   * ```ts
   * sheet.addConditionalFormatting({
   *   ref: 'A1:A100',
   *   rules: [{
   *     type: 'colorScale',
   *     cfvo: [
   *       { type: 'min' },
   *       { type: 'max' },
   *     ],
   *     color: [
   *       { argb: 'FFFF0000' },
   *       { argb: 'FF00FF00' },
   *     ],
   *   }],
   * });
   * ```
   */
  addConditionalFormatting(cf: unknown): void {
    this.conditionalFormattings.push(cf);
  }

  /**
   * Removes conditional formatting rules.
   *
   * @param filter - Index to remove, filter function, or undefined to remove all
   */
  removeConditionalFormatting(filter: unknown): void {
    if (typeof filter === 'number') {
      this.conditionalFormattings.splice(filter, 1);
    } else if (filter instanceof Function) {
      this.conditionalFormattings = this.conditionalFormattings.filter(filter);
    } else {
      this.conditionalFormattings = [];
    }
  }

  // ===========================================================================
  // Deprecated
  get tabColor(): unknown {
    // eslint-disable-next-line no-console
    console.trace('worksheet.tabColor property is now deprecated. Please use worksheet.properties.tabColor');
    return this.properties.tabColor;
  }

  set tabColor(value: unknown) {
    // eslint-disable-next-line no-console
    console.trace('worksheet.tabColor property is now deprecated. Please use worksheet.properties.tabColor');
    this.properties.tabColor = value;
  }

  // ===========================================================================
  // Model

  get model(): WorksheetModel {
    const model: WorksheetModel = {
      id: this.id,
      name: this.name,
      dataValidations: this.dataValidations.model,
      properties: this.properties,
      state: this.state,
      pageSetup: this.pageSetup,
      headerFooter: this.headerFooter,
      rowBreaks: this.rowBreaks,
      views: this.views,
      autoFilter: this.autoFilter,
      media: this._media.map(medium => medium.model),
      sheetProtection: this.sheetProtection,
      tables: Object.values(this.tables).map(table => table.model),
      pivotTables: this.pivotTables,
      conditionalFormattings: this.conditionalFormattings,
    };

    // =================================================
    // columns
    model.cols = Column.toModel(this.columns);

    // ==========================================================
    // Rows
    const rows: unknown[] = [];
    const dimensions = new Range();
    this._rows.forEach(row => {
      const rowModel = row && row.model;
      if (rowModel) {
        dimensions.expand(rowModel.number as number, rowModel.min as number, rowModel.number as number, rowModel.max as number);
        rows.push(rowModel);
      }
    });
    model.rows = rows;
    model.dimensions = dimensions;

    // ==========================================================
    // Merges
    model.merges = [];
    _.each(this._merges, merge => {
      model.merges.push(merge.range);
    });

    return model;
  }

  _parseRows(model: WorksheetModel): void {
    this._rows = [];
    (model.rows as unknown[]).forEach((rowModel: unknown) => {
      const rm = rowModel as Record<string, unknown>;
      const row = new Row(this, rm.number as number);
      this._rows[row.number - 1] = row;
      row.model = rowModel as Record<string, unknown>;
    });
  }

  _parseMergeCells(model: WorksheetModel): void {
    // Support both 'merges' (from model getter) and 'mergeCells' (from xform parsing)
    const merges = model.merges || (model as unknown as {mergeCells?: unknown[]}).mergeCells;
    _.each(merges, (merge: unknown) => {
      this.mergeCellsWithoutStyle(merge as string);
    });
  }

  set model(value: WorksheetModel) {
    this.name = value.name;
    this._columns = Column.fromModel(this, value.cols as unknown[]);
    this._parseRows(value);

    this._parseMergeCells(value);
    this.dataValidations = new DataValidations(value.dataValidations as Record<string, unknown> | undefined);
    this.properties = value.properties as Record<string, unknown>;
    this.pageSetup = value.pageSetup;
    this.headerFooter = value.headerFooter;
    this.views = value.views;
    this.autoFilter = value.autoFilter;
    this._media = (value.media as unknown[]).map(medium => new Image(this, medium));
    this.sheetProtection = value.sheetProtection;
    this.tables = ((value.tables as unknown[]).reduce((tables: Record<string, Table>, table: unknown) => {
      const t = new Table();
      (t as any).model = table;
      tables[(table as Record<string, unknown>).name as string] = t;
      return tables;
    }, {} as Record<string, Table>)) as Record<string, Table>;
    this.pivotTables = value.pivotTables;
    this.conditionalFormattings = value.conditionalFormattings;
  }
}

export default Worksheet;