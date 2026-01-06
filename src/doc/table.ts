/* eslint-disable max-classes-per-file */

import colCache from '../utils/col-cache.ts';
import type Worksheet from './worksheet.ts';
import type { CellValueType } from '../types/index.ts';

export interface TableColumnModel {
  name: string;
  filterButton?: boolean;
  style?: unknown;
  totalsRowLabel?: unknown;
  totalsRowFunction?: string;
  totalsRowResult?: unknown;
  totalsRowFormula?: unknown;
  [key: string]: unknown;
}

export interface TableModel {
  name: string;
  displayName?: string;
  ref?: string;
  headerRow?: boolean;
  totalsRow?: boolean;
  style?: Record<string, unknown>;
  columns: TableColumnModel[];
  rows: unknown[][];
  [key: string]: unknown;
}

/**
 * Table Column wrapper for column-level metadata and styling.
 */
class Column {
  table: Table;
  column: TableColumnModel;
  index: number;

  constructor(table: Table, column: TableColumnModel, index: number) {
    this.table = table;
    this.column = column;
    this.index = index;
  }

  _set(name: string, value: unknown): void {
    this.table.cacheState();
    this.column[name] = value;
  }

  /* eslint-disable lines-between-class-members */
  get name(): string {
    return this.column.name;
  }
  set name(value: string) {
    this._set('name', value);
  }

  get filterButton(): boolean {
    return this.column.filterButton;
  }
  set filterButton(value: boolean) {
    this.column.filterButton = value;
  }

  get style(): unknown {
    return this.column.style;
  }
  set style(value: unknown) {
    this.column.style = value;
  }

  get totalsRowLabel(): unknown {
    return this.column.totalsRowLabel;
  }
  set totalsRowLabel(value: unknown) {
    this._set('totalsRowLabel', value);
  }

  get totalsRowFunction(): unknown {
    return this.column.totalsRowFunction;
  }
  set totalsRowFunction(value: unknown) {
    this._set('totalsRowFunction', value);
  }

  get totalsRowResult(): unknown {
    return this.column.totalsRowResult;
  }
  set totalsRowResult(value: unknown) {
    this._set('totalsRowResult', value);
  }

  get totalsRowFormula(): unknown {
    return this.column.totalsRowFormula;
  }
  set totalsRowFormula(value: unknown) {
    this._set('totalsRowFormula', value);
  }
  /* eslint-enable lines-between-class-members */
}

/**
 * Table represents an Excel table with columns, rows, and filters.
 */
class Table {
  worksheet?: Worksheet;
  table?: TableModel;
  columns: Column[] = [];
  state?: string;
  _cache?: { ref: string; width: number; tableHeight: number };

  constructor(worksheet?: Worksheet, table?: TableModel) {
    this.worksheet = worksheet;
    if (table) {
      this.table = table;
      this.validate();
      this.store();
    }
  }

  getFormula(column: TableColumnModel): string | null {
    switch (column.totalsRowFunction) {
      case 'none':
        return null;
      case 'average':
        return `SUBTOTAL(101,${this.table.name}[${column.name}])`;
      case 'countNums':
        return `SUBTOTAL(102,${this.table.name}[${column.name}])`;
      case 'count':
        return `SUBTOTAL(103,${this.table.name}[${column.name}])`;
      case 'max':
        return `SUBTOTAL(104,${this.table.name}[${column.name}])`;
      case 'min':
        return `SUBTOTAL(105,${this.table.name}[${column.name}])`;
      case 'stdDev':
        return `SUBTOTAL(106,${this.table.name}[${column.name}])`;
      case 'var':
        return `SUBTOTAL(107,${this.table.name}[${column.name}])`;
      case 'sum':
        return `SUBTOTAL(109,${this.table.name}[${column.name}])`;
      case 'custom':
        return column.totalsRowFormula as string;
      default:
        throw new Error(`Invalid Totals Row Function: ${column.totalsRowFunction}`);
    }
  }

  get width(): number {
    // width of the table
    return this.table.columns.length;
  }

  get height(): number {
    // height of the table data
    return this.table.rows.length;
  }

  get filterHeight(): number {
    // height of the table data plus optional header row
    return this.height + (this.table.headerRow ? 1 : 0);
  }

  get tableHeight(): number {
    // full height of the table on the sheet
    return this.filterHeight + (this.table.totalsRow ? 1 : 0);
  }

  validate() {
    const {table} = this;
    // set defaults and check is valid
    const assign = (o, name, dflt) => {
      if (o[name] === undefined) {
        o[name] = dflt;
      }
    };
    assign(table, 'headerRow', true);
    assign(table, 'totalsRow', false);

    assign(table, 'style', {});
    assign(table.style, 'theme', 'TableStyleMedium2');
    assign(table.style, 'showFirstColumn', false);
    assign(table.style, 'showLastColumn', false);
    assign(table.style, 'showRowStripes', false);
    assign(table.style, 'showColumnStripes', false);

    const assert = (test, message) => {
      if (!test) {
        throw new Error(message);
      }
    };
    assert(table.ref, 'Table must have ref');
    assert(table.columns, 'Table must have column definitions');
    assert(table.rows, 'Table must have row definitions');

    table.tl = colCache.decodeAddress(table.ref);
    const {row, col} = table.tl as { row: number; col: number };
    assert(row > 0, 'Table must be on valid row');
    assert(col > 0, 'Table must be on valid col');

    const {width, filterHeight, tableHeight} = this;

    // autoFilterRef is a range that includes optional headers only
    table.autoFilterRef = colCache.encode(row, col, row + filterHeight - 1, col + width - 1);

    // tableRef is a range that includes optional headers and totals
    table.tableRef = colCache.encode(row, col, row + tableHeight - 1, col + width - 1);

    table.columns.forEach((column, i) => {
      assert(column.name, `Column ${i} must have a name`);
      if (i === 0) {
        assign(column, 'totalsRowLabel', 'Total');
      } else {
        assign(column, 'totalsRowFunction', 'none');
        column.totalsRowFormula = this.getFormula(column);
      }
    });
  }

  store() {
    // where the table needs to store table data, headers, footers in
    // the sheet...
    const assignStyle = (cell, style) => {
      if (style) {
        Object.keys(style).forEach(key => {
          cell.style[key] = style[key];
        });
      }
    };

    const {worksheet, table} = this;
    const {row, col} = table.tl as { row: number; col: number };
    let count = 0;
    if (table.headerRow) {
      const r = worksheet.getRow(row + count++);
      table.columns.forEach((column, j) => {
        const {style, name} = column;
        const cell = r.getCell(col + j);
        cell.value = name;
        assignStyle(cell, style);
      });
    }
    table.rows.forEach(data => {
      const r = worksheet.getRow(row + count++);
      data.forEach((value, j) => {
        const cell = r.getCell(col + j);
        cell.value = value as CellValueType;

        assignStyle(cell, table.columns[j].style);
      });
    });

    if (table.totalsRow) {
      const r = worksheet.getRow(row + count++);
      table.columns.forEach((column, j) => {
        const cell = r.getCell(col + j);
        if (j === 0) {
          cell.value = column.totalsRowLabel as CellValueType;
        } else {
          const formula = this.getFormula(column);
          if (formula) {
            cell.value = {
              formula: column.totalsRowFormula,
              result: column.totalsRowResult,
            } as CellValueType;
          } else {
            cell.value = null;
          }
        }

        assignStyle(cell, column.style);
      });
    }
  }

  load(worksheet: Worksheet) {
    // where the table will read necessary features from a loaded sheet
    const {table} = this;
    const {row, col} = table.tl as { row: number; col: number };
    let count = 0;
    if (table.headerRow) {
      const r = worksheet.getRow(row + count++);
      table.columns.forEach((column, j) => {
        const cell = r.getCell(col + j);
        cell.value = column.name as CellValueType;
      });
    }
    table.rows.forEach(data => {
      const r = worksheet.getRow(row + count++);
      data.forEach((value, j) => {
        const cell = r.getCell(col + j);
        cell.value = value as CellValueType;
      });
    });

    if (table.totalsRow) {
      const r = worksheet.getRow(row + count++);
      table.columns.forEach((column, j) => {
        const cell = r.getCell(col + j);
        if (j === 0) {
          cell.value = column.totalsRowLabel as CellValueType;
        } else {
          const formula = this.getFormula(column);
          if (formula) {
            cell.value = {
              formula: column.totalsRowFormula,
              result: column.totalsRowResult,
            } as CellValueType;
          }
        }
      });
    }
  }

  get model(): TableModel | undefined {
    return this.table;
  }

  set model(value: TableModel | undefined) {
    this.table = value;
  }

  // ================================================================
  // TODO: Mutating methods
  cacheState(): void {
    if (!this._cache) {
      this._cache = {
        ref: this.ref,
        width: this.width,
        tableHeight: this.tableHeight,
      };
    }
  }

  commit(): void {
    // changes may have been made that might have on-sheet effects
    if (!this._cache) {
      return;
    }

    // check things are ok first
    this.validate();

    const ref = colCache.decodeAddress(this._cache.ref);
    if (this.ref !== this._cache.ref) {
      // wipe out whole table footprint at previous location
      for (let i = 0; i < this._cache.tableHeight; i++) {
        const row = this.worksheet.getRow(ref.row + i);
        for (let j = 0; j < this._cache.width; j++) {
          const cell = row.getCell(ref.col + j);
          cell.value = null;
        }
      }
    } else {
      // clear out below table if it has shrunk
      for (let i = this.tableHeight; i < this._cache.tableHeight; i++) {
        const row = this.worksheet.getRow(ref.row + i);
        for (let j = 0; j < this._cache.width; j++) {
          const cell = row.getCell(ref.col + j);
          cell.value = null;
        }
      }

      // clear out to right of table if it has lost columns
      for (let i = 0; i < this.tableHeight; i++) {
        const row = this.worksheet.getRow(ref.row + i);
        for (let j = this.width; j < this._cache.width; j++) {
          const cell = row.getCell(ref.col + j);
          cell.value = null;
        }
      }
    }

    this.store();
  }

  addRow(values: unknown[], rowNumber?: number): void {
    // Add a row of data, either insert at rowNumber or append
    this.cacheState();

    if (rowNumber === undefined) {
      this.table.rows.push(values);
    } else {
      this.table.rows.splice(rowNumber, 0, values);
    }
  }

  removeRows(rowIndex: number, count = 1): void {
    // Remove a rows of data
    this.cacheState();
    this.table.rows.splice(rowIndex, count);
  }

  getColumn(colIndex: number): Column {
    const column = this.table.columns[colIndex];
    return new Column(this, column, colIndex);
  }

  addColumn(column: TableColumnModel, values: unknown[], colIndex?: number): void {
    this.cacheState();

    if (!this.table) return;
    if (colIndex === undefined) {
      this.table.columns.push(column);
      this.table.rows.forEach((row: unknown[], i: number) => {
        (row as unknown[]).push(values[i]);
      });
    } else {
      this.table.columns.splice(colIndex, 0, column);
      this.table.rows.forEach((row, i) => {
        row.splice(colIndex, 0, values[i]);
      });
    }
  }

  removeColumns(colIndex: number, count = 1): void {
    // Remove a column with data
    this.cacheState();

    this.table.columns.splice(colIndex, count);
    this.table.rows.forEach(row => {
      row.splice(colIndex, count);
    });
  }

  _assign(target: Record<string, unknown>, prop: string, value: unknown): void {
    this.cacheState();
    target[prop] = value;
  }

  /* eslint-disable lines-between-class-members */
  get ref(): string {
    return this.table.ref;
  }
  set ref(value: string) {
    this._assign(this.table, 'ref', value);
  }

  get name(): string {
    return this.table.name;
  }
  set name(value: string) {
    this.table.name = value;
  }

  get displayName(): string {
    return (this.table.displayName as string) || this.table.name;
  }
  set displayName(value: string) {
    this.table.displayName = value;
  }

  get headerRow(): boolean {
    return this.table.headerRow as boolean;
  }
  set headerRow(value: boolean) {
    this._assign(this.table, 'headerRow', value);
  }

  get totalsRow(): boolean {
    return this.table.totalsRow as boolean;
  }
  set totalsRow(value: boolean) {
    this._assign(this.table, 'totalsRow', value);
  }

  get theme(): string {
    return this.table.style?.name as string;
  }
  set theme(value: string) {
    this.table.style!.name = value;
  }

  get showFirstColumn(): boolean {
    return this.table.style?.showFirstColumn as boolean;
  }
  set showFirstColumn(value: boolean) {
    this.table.style!.showFirstColumn = value;
  }

  get showLastColumn(): boolean {
    return this.table.style?.showLastColumn as boolean;
  }
  set showLastColumn(value: boolean) {
    this.table.style!.showLastColumn = value;
  }

  get showRowStripes(): boolean {
    return this.table.style?.showRowStripes as boolean;
  }
  set showRowStripes(value: boolean) {
    this.table.style!.showRowStripes = value;
  }

  get showColumnStripes(): boolean {
    return this.table.style?.showColumnStripes as boolean;
  }
  set showColumnStripes(value: boolean) {
    this.table.style!.showColumnStripes = value;
  }
  /* eslint-enable lines-between-class-members */
}

export default Table;