
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
    if ((location as any).top) {
      for (let col = (location as any).left; col <= (location as any).right; col++) {
        for (let row = (location as any).top; row <= (location as any).bottom; row++) {
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

  _explore(matrix: CellMatrix, cell: Record<string, unknown>): Range {
    (cell as any).mark = false;
    const {sheetName} = cell;

    const range = new Range((cell as any).row, (cell as any).col, (cell as any).row, (cell as any).col, sheetName as string);
    let x;
    let y;

    // grow vertical - only one col to worry about
    function vGrow(yy: number, edge: string): boolean {
      const c = matrix.findCellAt(sheetName, yy, cell.col);
      if (!c || !c.mark) {
        return false;
      }
      range[edge] = yy;
      c.mark = false;
      return true;
    }
    for (y = cell.row - 1; vGrow(y, 'top'); y--);
    for (y = cell.row + 1; vGrow(y, 'bottom'); y++);

    // grow horizontal - ensure all rows can grow
    function hGrow(xx: number, edge: string): boolean {
      const cells: unknown[] = [];
      for (y = (range as any).top; y <= (range as any).bottom; y++) {
        const c = matrix.findCellAt(sheetName as string, y, (cell as any).col);
        if (c && (c as any).mark) {
          cells.push(c);
        } else {
          return false;
        }
      }
      (range as any)[edge] = xx;
      for (let i = 0; i < cells.length; i++) {
        ((cells[i] as any)).mark = false;
      }
      return true;
    }
    for (x = cell.col - 1; hGrow(x, 'left'); x--);
    for (x = cell.col + 1; hGrow(x, 'right'); x++);

    return range;
  }

  getRanges(name: string, matrix?: CellMatrix): Record<string, unknown> {
    const m = matrix || this.matrixMap[name];

    if (!m) {
      return {name, ranges: []};
    }

    m.forEach((cell: unknown) => {
      ((cell as any)).mark = true;
    });
    const ranges = matrix
      .map((cell: unknown) => ((cell as any).mark && this._explore(m, cell)))
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
        if ((cell as any).row !== row || (cell as any).col !== col) {
          (cell as any).row = row;
          (cell as any).col = col;
          (cell as any).address = colCache.n2l(col) + row;
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
      (definedName: unknown) => (definedName as any).ranges?.length
    );
  }

  set model(value: DefinedNameModel[]) {
    const matrixMap = (this.matrixMap = {});
    value.forEach((definedName: DefinedNameModel) => {
      const matrix = (matrixMap[(definedName as any).name] = new CellMatrix());
      ((definedName as any).ranges as string[]).forEach((rangeStr: string) => {
        if (rangeRegexp.test(rangeStr.split('!').pop() || '')) {
          matrix.addCell(rangeStr);
        }
      });
    });
  }
}

export default DefinedNames;