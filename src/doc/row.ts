
import _ from '../utils/under-dash.ts';
import Enums from './enums.ts';
import colCache from '../utils/col-cache.ts';
import Cell from './cell.ts';
import type Worksheet from './worksheet.ts';
import type { CellValueType } from '../types/index.ts';

/**
 * Represents a row in a worksheet.
 *
 * Rows provide access to cells and row-level formatting like height,
 * visibility, and outline levels.
 *
 * @example
 * ```ts
 * const row = sheet.getRow(1);
 *
 * // Set values
 * row.values = ['Name', 'Age', 'Email'];
 * row.values = { name: 'Alice', age: 30 };
 *
 * // Access cells
 * row.getCell(1).value = 'Hello';
 * row.getCell('A').value = 'Hello';
 *
 * // Style the row
 * row.height = 20;
 * row.font = { bold: true };
 * row.hidden = true;
 * ```
 */
class Row {
  /** Row height in points */
  height?: number;
  /** Row-level style applied to all cells */
  style: Record<string, unknown>;

  _worksheet: Worksheet;
  _number: number;
  _cells: (Cell | undefined)[];
  _outlineLevel: number;
  _hidden: boolean;

  constructor(worksheet: Worksheet, number: number) {
    this._worksheet = worksheet;
    this._number = number;
    this._cells = [];
    this.style = {};
    this._outlineLevel = 0;
    this._hidden = false;
  }

  /** Row number (1-based) */
  get number(): number {
    return this._number;
  }

  /** The parent worksheet */
  get worksheet(): Worksheet {
    return this._worksheet;
  }

  /**
   * Commits the row for streaming writes.
   * Informs the streaming writer that this row is complete.
   * Has no effect on regular Worksheet operations.
   */
  commit(): void {
    this._worksheet._commitRow(this); // eslint-disable-line no-underscore-dangle
  }

  /** @internal */
  destroy(): void {
    delete this._worksheet;
    delete this._cells;
    delete this.style;
  }

  /**
   * Finds a cell by column number. Returns undefined if the cell doesn't exist.
   *
   * @param colNumber - Column number (1-based)
   */
  findCell(colNumber: number): Cell | undefined {
    return this._cells[colNumber - 1];
  }

  /** @internal */
  getCellEx(address: {col: number; row: number; address: string}): Cell {
    let cell = this._cells[address.col - 1];
    if (!cell) {
      const column = this._worksheet.getColumn(address.col);
      cell = new Cell(this, column, address.address);
      this._cells[address.col - 1] = cell;
    }
    return cell;
  }

  /**
   * Gets a cell by column number, letter, or key. Creates the cell if needed.
   *
   * @param col - Column number (1-based), letter ('A', 'B'), or key name
   * @returns The cell
   *
   * @example
   * ```ts
   * row.getCell(1).value = 'First';
   * row.getCell('B').value = 'Second';
   * row.getCell('name').value = 'Alice'; // requires column key
   * ```
   */
  getCell(col: number | string): Cell {
    let colNum = col as number;
    if (typeof col === 'string') {
      // is it a key?
      const column = this._worksheet.getColumnKey(col);
      if (column) {
        colNum = column.number;
      } else {
        colNum = colCache.l2n(col);
      }
    }
    return (
      this._cells[colNum - 1] ||
      this.getCellEx({
        address: colCache.encodeAddress(this._number, colNum),
        row: this._number,
        col: colNum,
      })
    );
  }

  // remove cell(s) and shift all higher cells down by count
  splice(start: number, count: number, ...inserts: unknown[]): void {
    const nKeep = start + count;
    const nExpand = inserts.length - count;
    const nEnd = this._cells.length;
    let i;
    let cSrc;
    let cDst;

    if (nExpand < 0) {
      // remove cells
      for (i = start + inserts.length; i <= nEnd; i++) {
        cDst = this._cells[i - 1];
        cSrc = this._cells[i - nExpand - 1];
        if (cSrc) {
          cDst = this.getCell(i);
          cDst.value = cSrc.value;
          cDst.style = cSrc.style;
          // eslint-disable-next-line no-underscore-dangle
          cDst._comment = cSrc._comment;
        } else if (cDst) {
          cDst.value = null;
          cDst.style = {};
          // eslint-disable-next-line no-underscore-dangle
          cDst._comment = undefined;
        }
      }
    } else if (nExpand > 0) {
      // insert new cells
      for (i = nEnd; i >= nKeep; i--) {
        cSrc = this._cells[i - 1];
        if (cSrc) {
          cDst = this.getCell(i + nExpand);
          cDst.value = cSrc.value;
          cDst.style = cSrc.style;
          // eslint-disable-next-line no-underscore-dangle
          cDst._comment = cSrc._comment;
        } else {
          this._cells[i + nExpand - 1] = undefined;
        }
      }
    }

    // now add the new values
    for (i = 0; i < inserts.length; i++) {
      cDst = this.getCell(start + i);
      cDst.value = inserts[i];
      cDst.style = {};
      // eslint-disable-next-line no-underscore-dangle
      cDst._comment = undefined;
    }
  }

  /**
   * Iterates over cells in this row.
   *
   * @param options - Options with `includeEmpty: boolean`, or the iteratee
   * @param iteratee - Callback receiving (cell, colNumber)
   *
   * @example
   * ```ts
   * // Only cells with values
   * row.eachCell((cell, colNumber) => {
   *   console.log(`Column ${colNumber}: ${cell.value}`);
   * });
   *
   * // Include empty cells
   * row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
   *   console.log(`Column ${colNumber}`);
   * });
   * ```
   */
  eachCell(options: Record<string, unknown> | ((cell: Cell, colNum: number) => void), iteratee?: (cell: Cell, colNum: number) => void): void {
    let iter: ((cell: Cell, colNum: number) => void) | undefined = iteratee;
    let opts: Record<string, unknown> | null = null;
    if (typeof options === 'function') {
      iter = options;
      opts = null;
    } else {
      opts = options;
    }
    if (opts && opts.includeEmpty) {
      const n = this._cells.length;
      for (let i = 1; i <= n; i++) {
        (iter as Function)(this.getCell(i), i);
      }
    } else {
      this._cells.forEach((cell, index) => {
        if (cell && cell.type !== Enums.ValueType.Null) {
          (iter as Function)(cell, index + 1);
        }
      });
    }
  }

  // ===========================================================================
  // Page Breaks

  /**
   * Adds a page break after this row.
   *
   * @param lft - Left column boundary
   * @param rght - Right column boundary
   */
  addPageBreak(lft?: number, rght?: number): void {
    const ws = this._worksheet;
    const left = Math.max(0, (lft || 0) - 1) || 0;
    const right = Math.max(0, (rght || 0) - 1) || 16838;
    const pb: Record<string, number> = {
      id: this._number,
      max: right,
      man: 1,
    };
    if (left) pb.min = left;

    ws.rowBreaks.push(pb);
  }

  /**
   * Cell values as a sparse array (index = column number).
   *
   * @example
   * ```ts
   * const vals = row.values;
   * // vals[1] = 'A1 value', vals[2] = 'B1 value', etc.
   * ```
   */
  get values(): unknown[] {
    const values: unknown[] = [];
    this._cells.forEach(cell => {
      if (cell && cell.type !== Enums.ValueType.Null) {
        values[cell.col] = cell.value;
      }
    });
    return values;
  }

  /**
   * Sets cell values from an array or object.
   *
   * @example
   * ```ts
   * // Contiguous array (starting at column A)
   * row.values = ['Alice', 30, 'alice@example.com'];
   *
   * // Sparse array (index = column number)
   * row.values = [, 'A1', , 'C1']; // columns A and C
   *
   * // Object with column keys
   * row.values = { name: 'Alice', age: 30 };
   * ```
   */
  set values(value: unknown[] | Record<string, unknown> | undefined) {
    // this operation is not additive - any prior cells are removed
    this._cells = [];
    if (!value) {
      // empty row
    } else if (value instanceof Array) {
      let offset = 0;
      if (value.hasOwnProperty('0')) {
        // contiguous array - start at column 1
        offset = 1;
      }
      value.forEach((item, index) => {
        if (item !== undefined) {
          this.getCellEx({
            address: colCache.encodeAddress(this._number, index + offset),
            row: this._number,
            col: index + offset,
          }).value = item as CellValueType;
        }
      });
    } else {
      // assume object with column keys
      this._worksheet.eachColumnKey((column, key) => {
        if (value[key] !== undefined) {
          this.getCellEx({
            address: colCache.encodeAddress(this._number, column.number),
            row: this._number,
            col: column.number,
          }).value = value[key] as CellValueType;
        }
      });
    }
  }

  /** Whether this row has any cells with values */
  get hasValues(): boolean {
    return _.some(this._cells, cell => cell && cell.type !== Enums.ValueType.Null);
  }

  /** Total number of cells in this row (including empty) */
  get cellCount(): number {
    return this._cells.length;
  }

  /** Number of cells with actual values */
  get actualCellCount(): number {
    let count = 0;
    this.eachCell(() => {
      count++;
    });
    return count;
  }

  /**
   * The column range of cells with values.
   * Returns `{ min, max }` or null if row is empty.
   */
  get dimensions(): { min: number; max: number } | null {
    let min = 0;
    let max = 0;
    this._cells.forEach(cell => {
      if (cell && cell.type !== Enums.ValueType.Null) {
        if (!min || min > cell.col) {
          min = cell.col;
        }
        if (max < cell.col) {
          max = cell.col;
        }
      }
    });
    return min > 0
      ? {
          min,
          max,
        }
      : null;
  }

  // =========================================================================
  // Styles

  /** @internal */
  _applyStyle(name: string, value: unknown): unknown {
    this.style[name] = value;
    this._cells.forEach(cell => {
      if (cell) {
        cell[name] = value;
      }
    });
    return value;
  }

  /** Number format applied to all cells in this row */
  get numFmt(): unknown {
    return this.style.numFmt;
  }

  set numFmt(value: unknown) {
    this._applyStyle('numFmt', value);
  }

  /** Font style applied to all cells in this row */
  get font(): unknown {
    return this.style.font;
  }

  set font(value: unknown) {
    this._applyStyle('font', value);
  }

  /** Alignment applied to all cells in this row */
  get alignment(): unknown {
    return this.style.alignment;
  }

  set alignment(value: unknown) {
    this._applyStyle('alignment', value);
  }

  /** Protection settings applied to all cells in this row */
  get protection(): unknown {
    return this.style.protection;
  }

  set protection(value: unknown) {
    this._applyStyle('protection', value);
  }

  /** Border style applied to all cells in this row */
  get border(): unknown {
    return this.style.border;
  }

  set border(value: unknown) {
    this._applyStyle('border', value);
  }

  /** Fill style applied to all cells in this row */
  get fill(): unknown {
    return this.style.fill;
  }

  set fill(value: unknown) {
    this._applyStyle('fill', value);
  }

  /** Whether this row is hidden */
  get hidden(): boolean {
    return !!this._hidden;
  }

  set hidden(value: boolean) {
    this._hidden = value;
  }

  /** Outline/grouping level (0-7) */
  get outlineLevel(): number {
    return this._outlineLevel || 0;
  }

  set outlineLevel(value: number) {
    this._outlineLevel = value;
  }

  /** Whether this row is collapsed in outline */
  get collapsed(): boolean {
    return !!(
      this._outlineLevel && this._outlineLevel >= (this._worksheet.properties.outlineLevelRow as number)
    );
  }

  // =========================================================================
  get model(): Record<string, unknown> | null {
    const cells: unknown[] = [];
    let min = 0;
    let max = 0;
    this._cells.forEach(cell => {
      if (cell) {
        const cellModel = cell.model;
        if (cellModel) {
          if (!min || min > cell.col) {
            min = cell.col;
          }
          if (max < cell.col) {
            max = cell.col;
          }
          cells.push(cellModel);
        }
      }
    });

    return this.height || cells.length
      ? {
          cells,
          number: this.number,
          min,
          max,
          height: this.height,
          style: this.style,
          hidden: this.hidden,
          outlineLevel: this.outlineLevel,
          collapsed: this.collapsed,
        }
      : null;
  }

  set model(value: Record<string, unknown>) {
    if ((value.number as number) !== this._number) {
      throw new Error('Invalid row number in model');
    }
    this._cells = [];
    let previousAddress: { row: number; col: number; address: string; $col$row: string } | null = null;
    const cells = value.cells as Array<{ type?: number; address?: string; [key: string]: unknown }>;
    cells.forEach(cellModel => {
      switch (cellModel.type) {
        case Cell.Types.Merge:
          // special case - don't add this types
          break;
        default: {
          let address: { row: number; col: number; address: string; $col$row: string } | undefined;
          if (cellModel.address) {
            address = colCache.decodeAddress(cellModel.address);
          } else if (previousAddress) {
            // This is a <c> element without an r attribute
            // Assume that it's the cell for the next column
            const {row} = previousAddress;
            const col = previousAddress.col + 1;
            address = {
              row,
              col,
              address: colCache.encodeAddress(row, col),
              $col$row: `$${colCache.n2l(col)}$${row}`,
            };
          }
          previousAddress = address ?? null;
          const cell = this.getCellEx(address);
          cell.model = cellModel as Record<string, unknown>;
          break;
        }
      }
    });

    if (value.height) {
      this.height = value.height as number;
    } else {
      delete this.height;
    }

    this.hidden = value.hidden as boolean;
    this.outlineLevel = (value.outlineLevel as number) || 0;

    this.style = (value.style && JSON.parse(JSON.stringify(value.style))) || {};
  }
}

export default Row;