
import _ from '../utils/under-dash.ts';
import colCache from '../utils/col-cache.ts';
import CellMatrix from '../utils/cell-matrix.ts';
import Range from './range.ts';

const rangeRegexp = /[$](\w+)[$](\d+)(:[$](\w+)[$](\d+))?/;

export interface CellAddress {
  sheetName?: string;
  address?: string;
  row?: number;
  col?: number;
  top?: number;
  bottom?: number;
  left?: number;
  right?: number;
  mark?: boolean;
}

export interface DefinedNameModel {
  name: string;
  ranges: unknown[];
  [key: string]: unknown;
}

/**
 * DefinedNames tracks named ranges and cell references in a workbook.
 */
class DefinedNames {
  matrixMap: Record<string, CellMatrix> = {};

  constructor() {
    this.matrixMap = {};
  }

  getMatrix(name: string): CellMatrix {
    const matrix = this.matrixMap[name] || (this.matrixMap[name] = new CellMatrix());
    return matrix;
  }

  add(locStr: string, name: unknown): void {
    const location = colCache.decodeEx(locStr);
    this.addEx(location, name);
  }

  addEx(location: CellAddress, name: unknown): void {
    const matrix = this.getMatrix(name as string);
    if (location.top) {
      for (let col = location.left!; col <= location.right!; col++) {
        for (let row = location.top; row <= location.bottom!; row++) {
          const address: CellAddress = {
            sheetName: location.sheetName,
            address: colCache.n2l(col) + row,
            row,
            col,
          };

          matrix.addCellEx(address);
        }
      }
    } else {
      matrix.addCellEx(location);
    }
  }

  remove(locStr: string, name: unknown): void {
    const location = colCache.decodeEx(locStr);
    this.removeEx(location, name);
  }

  removeEx(location: CellAddress, name: unknown): void {
    const matrix = this.getMatrix(name as string);
    matrix.removeCellEx(location);
  }

  removeAllNames(location: CellAddress): void {
    _.each(this.matrixMap, (matrix: CellMatrix) => {
      matrix.removeCellEx(location);
    });
  }

  forEach(callback: (name: string, cell: unknown) => void): void {
    _.each(this.matrixMap, (matrix: CellMatrix, name: string) => {
      matrix.forEach((cell: unknown) => {
        callback(name, cell);
      });
    });
  }

  getNames(addressStr: string): unknown[] {
    return this.getNamesEx(colCache.decodeEx(addressStr));
  }

  getNamesEx(address: CellAddress): unknown[] {
    return _.map(this.matrixMap, (matrix: CellMatrix, name: string) => matrix.findCellEx(address) && name).filter(
      Boolean
    );
  }

  _explore(matrix: CellMatrix, cell: CellAddress): Range {
    cell.mark = false;
    const {sheetName} = cell;

    const range = new Range(cell.row!, cell.col!, cell.row!, cell.col!, sheetName as string);
    let x;
    let y;

    // grow vertical - only one col to worry about
    function vGrow(yy: number, edge: string): boolean {
      const c = matrix.findCellAt(sheetName, yy, cell.col!);
      if (!c || !c.mark) {
        return false;
      }
      (c as CellAddress).mark = false;
      range[edge as keyof typeof range] = yy as never;
      return true;
    }
    for (y = cell.row! - 1; vGrow(y, 'top'); y--);
    for (y = cell.row! + 1; vGrow(y, 'bottom'); y++);

    // grow horizontal - ensure all rows can grow
    function hGrow(xx: number, edge: string): boolean {
      const cells: CellAddress[] = [];
      for (y = range.top!; y <= range.bottom!; y++) {
        const c = matrix.findCellAt(sheetName as string, y, xx);
        if (c && (c as CellAddress).mark) {
          cells.push(c as CellAddress);
        } else {
          return false;
        }
      }
      range[edge as keyof typeof range] = xx as never;
      for (let i = 0; i < cells.length; i++) {
        cells[i].mark = false;
      }
      return true;
    }
    for (x = cell.col! - 1; hGrow(x, 'left'); x--);
    for (x = cell.col! + 1; hGrow(x, 'right'); x++);

    return range;
  }

  getRanges(name: string, matrix?: CellMatrix): Record<string, unknown> {
    const m = matrix || this.matrixMap[name];

    if (!m) {
      return {name, ranges: []};
    }

    m.forEach((cell: unknown) => {
      (cell as CellAddress).mark = true;
    });
    const ranges = m
      .map((cell: unknown) => {
        const cellAddr = cell as CellAddress;
        return cellAddr.mark && this._explore(m, cellAddr);
      })
      .filter(Boolean)
      .map((range: Range) => range.$shortRange);

    return {
      name,
      ranges,
    };
  }

  normaliseMatrix(matrix: CellMatrix, sheetName: string): void {
    matrix.forEachInSheet(sheetName, (cell: unknown, row: number, col: number) => {
      if (cell) {
        const cellAddr = cell as CellAddress;
        if (cellAddr.row !== row || cellAddr.col !== col) {
          cellAddr.row = row;
          cellAddr.col = col;
          cellAddr.address = colCache.n2l(col) + row;
        }
      }
    });
  }

  spliceRows(sheetName: string, start: number, numDelete: number, numInsert: number): void {
    _.each(this.matrixMap, matrix => {
      matrix.spliceRows(sheetName, start, numDelete, numInsert);
      this.normaliseMatrix(matrix, sheetName);
    });
  }

  spliceColumns(sheetName: string, start: number, numDelete: number, numInsert: number): void {
    _.each(this.matrixMap, matrix => {
      matrix.spliceColumns(sheetName, start, numDelete, numInsert);
      this.normaliseMatrix(matrix, sheetName);
    });
  }

  get model(): unknown[] {
    // To get names per cell - just iterate over all names finding cells if they exist
    return _.map(this.matrixMap, (matrix, name) => this.getRanges(name, matrix)).filter(
      (definedName: unknown) => {
        const model = definedName as DefinedNameModel;
        return model.ranges?.length;
      }
    );
  }

  set model(value: DefinedNameModel[]) {
    const matrixMap = (this.matrixMap = {});
    value.forEach((definedName: DefinedNameModel) => {
      const matrix = (matrixMap[definedName.name] = new CellMatrix());
      (definedName.ranges as string[]).forEach((rangeStr: string) => {
        if (rangeRegexp.test(rangeStr.split('!').pop() || '')) {
          matrix.addCell(rangeStr);
        }
      });
    });
  }
}

export default DefinedNames;