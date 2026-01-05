
import _ from '../utils/under-dash.ts';
import Enums from './enums.ts';
import colCache from '../utils/col-cache.ts';
import type Worksheet from './worksheet.ts';

const DEFAULT_COLUMN_WIDTH = 9;

/**
 * Represents a column in a worksheet.
 *
 * Columns define properties like width, header text, and styles that
 * apply to all cells in the column.
 *
 * @example
 * ```ts
 * const col = sheet.getColumn('A');
 *
 * // Set properties
 * col.width = 20;
 * col.hidden = true;
 * col.header = 'Name';
 * col.key = 'name';
 *
 * // Apply styles to entire column
 * col.font = { bold: true };
 * col.alignment = { horizontal: 'center' };
 * ```
 */
class Column {
  /** Column width in characters */
  width?: number;
  /** Column-level style applied to all cells */
  style: Record<string, unknown>;
  
  /** @internal */
  _worksheet: Worksheet;
  /** @internal */
  _number: number;

  constructor(worksheet: Worksheet, number: number, defn?: Record<string, unknown>) {
    this._worksheet = worksheet;
    this._number = number;
    if (defn !== false) {
      this.defn = defn;
    }
  }

  get number(): number {
    return this._number;
  }
  get worksheet(): Worksheet {
    return this._worksheet;
  }

  /** Column letter ('A', 'B', 'AA', etc.) */
  get letter(): string {
    return colCache.n2l(this._number);
  }

  /** Whether this column has a custom width (not default 9) */
  get isCustomWidth(): boolean {
    return this.width !== undefined && this.width !== DEFAULT_COLUMN_WIDTH;
  }

  /**
   * Column definition object with all properties.
   */
  get defn(): Record<string, unknown> {
    return {
      header: this._header,
      key: this.key,
      width: this.width,
      style: this.style,
      hidden: this.hidden,
      outlineLevel: this.outlineLevel,
    };
  }

  set defn(value: Record<string, unknown> | undefined): void {
    if (value) {
      this.key = value.key;
      this.width = value.width !== undefined ? (value.width as number) : DEFAULT_COLUMN_WIDTH;
      this.outlineLevel = value.outlineLevel as number;
      if (value.style) {
        this.style = value.style as Record<string, unknown>;
      } else {
        this.style = {};
      }

      // headers must be set after style
      this.header = value.header;
      this._hidden = !!value.hidden;
    } else {
      delete this._header;
      delete this._key;
      delete this.width;
      this.style = {};
      this.outlineLevel = 0;
    }
  }

  /** Header values as an array (for multi-row headers) */
  get headers(): unknown[] {
    return this._header && this._header instanceof Array ? this._header : [this._header];
  }

  /**
   * Column header text.
   * Can be a single value or array for multi-row headers.
   *
   * @example
   * ```ts
   * col.header = 'Name';
   * col.header = ['First Name', 'Given Name']; // multi-row
   * ```
   */
  get header(): unknown {
    return this._header;
  }

  set header(value: unknown): void {
    if (value !== undefined) {
      this._header = value;
      this.headers.forEach((text, index) => {
        this._worksheet.getCell(index + 1, this.number).value = text;
      });
    } else {
      this._header = undefined;
    }
  }

  /**
   * Column key for use with object-based row values.
   *
   * @example
   * ```ts
   * col.key = 'name';
   * // Now can use: row.values = { name: 'Alice' };
   * // Or: row.getCell('name').value = 'Alice';
   * ```
   */
  get key(): string | number | undefined {
    return this._key;
  }

  set key(value: string | number | undefined): void {
    const column = this._key && this._worksheet.getColumnKey(this._key);
    if (column === this) {
      this._worksheet.deleteColumnKey(this._key);
    }

    this._key = value;
    if (value) {
      this._worksheet.setColumnKey(this._key, this);
    }
  }

  /** Whether this column is hidden */
  get hidden(): boolean {
    return !!this._hidden;
  }

  set hidden(value: boolean): void {
    this._hidden = value;
  }

  /** Outline/grouping level (0-7) */
  get outlineLevel(): number {
    return this._outlineLevel || 0;
  }

  set outlineLevel(value: number): void {
    this._outlineLevel = value;
  }

  /** Whether this column is collapsed in outline */
  get collapsed(): boolean {
    return !!(
      this._outlineLevel && this._outlineLevel >= this._worksheet.properties.outlineLevelCol
    );
  }

  toString(): string {
    return JSON.stringify({
      key: this.key,
      width: this.width,
      headers: this.headers.length ? this.headers : undefined,
    });
  }

  equivalentTo(other: Column): boolean {
    return (
      this.width === other.width &&
      this.hidden === other.hidden &&
      this.outlineLevel === other.outlineLevel &&
      _.isEqual(this.style, other.style)
    );
  }

  get isDefault(): boolean {
    if (this.isCustomWidth) {
      return false;
    }
    if (this.hidden) {
      return false;
    }
    if (this.outlineLevel) {
      return false;
    }
    const s = this.style;
    if (s && (s.font || s.numFmt || s.alignment || s.border || s.fill || s.protection)) {
      return false;
    }
    return true;
  }

  /** Number of header rows */
  get headerCount(): number {
    return this.headers.length;
  }

  /**
   * Iterates over cells in this column.
   *
   * @param options - Options with `includeEmpty: boolean`, or the iteratee
   * @param iteratee - Callback receiving (cell, rowNumber)
   *
   * @example
   * ```ts
   * column.eachCell((cell, rowNumber) => {
   *   console.log(`Row ${rowNumber}: ${cell.value}`);
   * });
   * ```
   */
  eachCell(options?: Record<string, unknown> | ((cell: unknown, rowNum: number) => void), iteratee?: (cell: unknown, rowNum: number) => void): void {
    let iter = iteratee;
    let opts: Record<string, unknown> | undefined = undefined;
    const colNumber = this.number;
    if (!iter && typeof options === 'function') {
      iter = options;
    } else {
      opts = options as Record<string, unknown>;
    }
    const worksheetRecord = this._worksheet as Record<string, unknown>;
    (worksheetRecord.eachRow as Function)(opts, (row: unknown, rowNumber: number) => {
      const rowRecord = row as Record<string, unknown>;
      (iter as Function)((rowRecord.getCell as Function)(colNumber), rowNumber);
    });
  }

  /**
   * Cell values as a sparse array (index = row number).
   */
  get values(): unknown[] {
    const v: unknown[] = [];
    this.eachCell((cell, rowNumber) => {
      if (cell && cell.type !== Enums.ValueType.Null) {
        v[rowNumber] = cell.value;
      }
    });
    return v;
  }

  /**
   * Sets cell values for the entire column.
   *
   * @example
   * ```ts
   * // Contiguous array (starting at row 1)
   * column.values = ['Header', 'Value1', 'Value2'];
   *
   * // Sparse array (index = row number)
   * column.values = [, 'Row1', , 'Row3'];
   * ```
   */
  set values(v: unknown[]): void {
    if (!v) {
      return;
    }
    const colNumber = this.number;
    let offset = 0;
    if (v.hasOwnProperty('0')) {
      // assume contiguous array, start at row 1
      offset = 1;
    }
    v.forEach((value, index) => {
      this._worksheet.getCell(index + offset, colNumber).value = value;
    });
  }

  // =========================================================================
  // Styles

  /** @internal */
  _applyStyle(name: string, value: unknown): unknown {
    this.style[name] = value;
    this.eachCell(cell => {
      cell[name] = value;
    });
    return value;
  }

  /** Number format applied to all cells in this column */
  get numFmt(): unknown {
    return this.style.numFmt;
  }

  set numFmt(value: unknown) {
    this._applyStyle('numFmt', value);
  }

  /** Font style applied to all cells in this column */
  get font(): unknown {
    return this.style.font;
  }

  set font(value: unknown) {
    this._applyStyle('font', value);
  }

  /** Alignment applied to all cells in this column */
  get alignment(): unknown {
    return this.style.alignment;
  }

  set alignment(value: unknown) {
    this._applyStyle('alignment', value);
  }

  /** Protection settings applied to all cells in this column */
  get protection(): unknown {
    return this.style.protection;
  }

  set protection(value: unknown) {
    this._applyStyle('protection', value);
  }

  /** Border style applied to all cells in this column */
  get border(): unknown {
    return this.style.border;
  }

  set border(value: unknown) {
    this._applyStyle('border', value);
  }

  /** Fill style applied to all cells in this column */
  get fill(): unknown {
    return this.style.fill;
  }

  set fill(value: unknown) {
    this._applyStyle('fill', value);
  }

  // =============================================================================
  // static functions

  static toModel(columns: Column[] | undefined): Record<string, unknown>[] | undefined {
    // Convert array of Column into compressed list cols
    const cols: Record<string, unknown>[] = [];
    let col: Record<string, unknown> | null = null;
    if (columns) {
      columns.forEach((column, index) => {
        if (column.isDefault) {
          if (col) {
            col = null;
          }
        } else if (!col || !column.equivalentTo(col as Column)) {
          col = {
            min: index + 1,
            max: index + 1,
            width: column.width !== undefined ? column.width : DEFAULT_COLUMN_WIDTH,
            style: column.style,
            isCustomWidth: column.isCustomWidth,
            hidden: column.hidden,
            outlineLevel: column.outlineLevel,
            collapsed: column.collapsed,
          };
          cols.push(col);
        } else {
          col.max = index + 1;
        }
      });
    }
    return cols.length ? cols : undefined;
  }

  static fromModel(worksheet: Worksheet, cols: unknown[] | undefined): Column[] | null {
    cols = cols || [];
    const columns = [];
    let count = 1;
    let index = 0;
    /**
     * sort cols by min
     * If it is not sorted, the subsequent column configuration will be overwritten
     * */
    cols = cols.sort(function(pre, next)  {
      return pre.min - next.min;
    });
    while (index < cols.length) {
      const col = cols[index++];
      while (count < col.min) {
        columns.push(new Column(worksheet, count++));
      }
      while (count <= col.max) {
        columns.push(new Column(worksheet, count++, col));
      }
    }
    return columns.length ? columns : null;
  }
}

export default Column;