
import _ from '../utils/under-dash.ts';
import colCache from '../utils/col-cache.ts';
import CellMatrix from '../utils/cell-matrix.ts';
import Range from './range.ts';

const rangeRegexp = /[$](\w+)[$](\d+)(:[$](\w+)[$](\d+))?/;

/**
 * DefinedNames tracks named ranges and cell references in a workbook.
 */
class DefinedNames {
  constructor() {
    this.matrixMap = {};
  }

  getMatrix(name: string): any {
    const matrix = this.matrixMap[name] || (this.matrixMap[name] = new CellMatrix());
    return matrix;
  }

  // add a name to a cell. locStr in the form SheetName!$col$row or SheetName!$c1$r1:$c2:$r2
  add(locStr: string, name: unknown): void {
    const location = colCache.decodeEx(locStr);
    this.addEx(location, name);
  }

  addEx(location: any, name: unknown): void {
    const matrix = this.getMatrix(name as string);
    if (location.top) {
      for (let col = location.left; col <= location.right; col++) {
        for (let row = location.top; row <= location.bottom; row++) {
          const address = {
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

  removeEx(location: any, name: unknown): void {
    const matrix = this.getMatrix(name as string);
    matrix.removeCellEx(location);
  }

  removeAllNames(location: any): void {
    _.each(this.matrixMap, matrix => {
      matrix.removeCellEx(location);
    });
  }

  forEach(callback: (name: string, cell: any) => void): void {
    _.each(this.matrixMap, (matrix, name) => {
      matrix.forEach(cell => {
        callback(name, cell);
      });
    });
  }

  // get all the names of a cell
  getNames(addressStr: string): unknown[] {
    return this.getNamesEx(colCache.decodeEx(addressStr));
  }

  getNamesEx(address: any): unknown[] {
    return _.map(this.matrixMap, (matrix, name) => matrix.findCellEx(address) && name).filter(
      Boolean
    );
  }

  _explore(matrix: any, cell: any): Range {
    cell.mark = false;
    const {sheetName} = cell;

    const range = new Range(cell.row, cell.col, cell.row, cell.col, sheetName);
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
      const cells: any[] = [];
      for (y = range.top; y <= range.bottom; y++) {
        const c = matrix.findCellAt(sheetName, y, xx);
        if (c && c.mark) {
          cells.push(c);
        } else {
          return false;
        }
      }
      range[edge] = xx;
      for (let i = 0; i < cells.length; i++) {
        cells[i].mark = false;
      }
      return true;
    }
    for (x = cell.col - 1; hGrow(x, 'left'); x--);
    for (x = cell.col + 1; hGrow(x, 'right'); x++);

    return range;
  }

  getRanges(name: string, matrix?: any): Record<string, unknown> {
    matrix = matrix || this.matrixMap[name];

    if (!matrix) {
      return {name, ranges: []};
    }

    // mark and sweep!
    matrix.forEach((cell: any) => {
      cell.mark = true;
    });
    const ranges = matrix
      .map((cell: any) => cell.mark && this._explore(matrix, cell))
      .filter(Boolean)
      .map((range: Range) => range.$shortRange);

    return {
      name,
      ranges,
    };
  }

  normaliseMatrix(matrix: any, sheetName: string): void {
    // some of the cells might have shifted on specified sheet
    // need to reassign rows, cols
    matrix.forEachInSheet(sheetName, (cell: any, row: number, col: number) => {
      if (cell) {
        if (cell.row !== row || cell.col !== col) {
          cell.row = row;
          cell.col = col;
          cell.address = colCache.n2l(col) + row;
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
      (definedName: any) => definedName.ranges.length
    );
  }

  set model(value: any[]) {
    // value is [ { name, ranges }, ... ]
    const matrixMap = (this.matrixMap = {});
    value.forEach((definedName: any) => {
      const matrix = (matrixMap[definedName.name] = new CellMatrix());
      definedName.ranges.forEach((rangeStr: string) => {
        if (rangeRegexp.test(rangeStr.split('!').pop() || '')) {
          matrix.addCell(rangeStr);
        }
      });
    });
  }
}

export default DefinedNames;