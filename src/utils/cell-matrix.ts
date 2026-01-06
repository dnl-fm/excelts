
import _ from './under-dash.ts';
import colCache from './col-cache.ts';

interface CellAddress {
  sheetName: string;
  row: number;
  col: number;
  top?: number;
  bottom?: number;
  left?: number;
  right?: number;
}

interface CellData {
  sheetName: string;
  address: string;
  row: number;
  col: number;
  [key: string]: unknown;
}

type SheetData = (CellData | undefined)[][];

class CellMatrix {
  template?: unknown;
  sheets: Record<string, SheetData>;

  constructor(template?: unknown) {
    this.template = template;
    this.sheets = {};
  }

  addCell(addressStr: string): void {
    this.addCellEx(colCache.decodeEx(addressStr));
  }

  getCell(addressStr: string): CellData | undefined {
    return this.findCellEx(colCache.decodeEx(addressStr), true);
  }

  findCell(addressStr: string): CellData | undefined {
    return this.findCellEx(colCache.decodeEx(addressStr), false);
  }

  findCellAt(sheetName: string, rowNumber: number, colNumber: number): CellData | undefined {
    const sheet = this.sheets[sheetName];
    const row = sheet && sheet[rowNumber];
    return row && row[colNumber];
  }

  addCellEx(address: CellAddress): void {
    if (address.top) {
      for (let row = address.top; row <= address.bottom!; row++) {
        for (let col = address.left!; col <= address.right!; col++) {
          this.getCellAt(address.sheetName, row, col);
        }
      }
    } else {
      this.findCellEx(address, true);
    }
  }

  getCellEx(address: CellAddress): CellData | undefined {
    return this.findCellEx(address, true);
  }

  findCellEx(address: CellAddress, create?: boolean): CellData | undefined {
    const sheet = this.findSheet(address, create);
    const row = this.findSheetRow(sheet, address, create);
    return this.findRowCell(row, address, create);
  }

  getCellAt(sheetName: string, rowNumber: number, colNumber: number): CellData {
    const sheet = this.sheets[sheetName] || (this.sheets[sheetName] = []);
    const row = sheet[rowNumber] || (sheet[rowNumber] = []);
    const cell =
      row[colNumber] ||
      (row[colNumber] = {
        sheetName,
        address: colCache.n2l(colNumber) + rowNumber,
        row: rowNumber,
        col: colNumber,
      });
    return cell;
  }

  removeCellEx(address: CellAddress): void {
    const sheet = this.findSheet(address);
    if (!sheet) {
      return;
    }
    const row = this.findSheetRow(sheet, address);
    if (!row) {
      return;
    }
    delete row[address.col];
  }

  forEachInSheet(sheetName: string, callback: (cell: CellData, rowNumber: number, colNumber: number) => void): void {
    const sheet = this.sheets[sheetName];
    if (sheet) {
      sheet.forEach((row, rowNumber) => {
        if (row) {
          row.forEach((cell, colNumber) => {
            if (cell) {
              callback(cell, rowNumber, colNumber);
            }
          });
        }
      });
    }
  }

  forEach(callback: (cell: CellData, rowNumber: number, colNumber: number) => void): void {
    _.each(this.sheets, (sheet, sheetName) => {
      this.forEachInSheet(sheetName, callback);
    });
  }

  map<T>(callback: (cell: CellData) => T): T[] {
    const results: T[] = [];
    this.forEach(cell => {
      results.push(callback(cell));
    });
    return results;
  }

  findSheet(address: CellAddress, create?: boolean): SheetData | undefined {
    const name = address.sheetName;
    if (this.sheets[name]) {
      return this.sheets[name];
    }
    if (create) {
      return (this.sheets[name] = []);
    }
    return undefined;
  }

  findSheetRow(sheet: SheetData | undefined, address: CellAddress, create?: boolean): (CellData | undefined)[] | undefined {
    const {row} = address;
    if (sheet && sheet[row]) {
      return sheet[row];
    }
    if (create && sheet) {
      return (sheet[row] = []);
    }
    return undefined;
  }

  findRowCell(row: (CellData | undefined)[] | undefined, address: CellAddress, create?: boolean): CellData | undefined {
    const {col} = address;
    if (row && row[col]) {
      return row[col];
    }
    if (create && row) {
      return (row[col] = this.template
        ? Object.assign(address, JSON.parse(JSON.stringify(this.template)))
        : address) as CellData;
    }
    return undefined;
  }

  spliceRows(sheetName: string, start: number, numDelete: number, numInsert: number): void {
    const sheet = this.sheets[sheetName];
    if (sheet) {
      const inserts: (CellData | undefined)[][] = [];
      for (let i = 0; i < numInsert; i++) {
        inserts.push([]);
      }
      sheet.splice(start, numDelete, ...inserts);
    }
  }

  spliceColumns(sheetName: string, start: number, numDelete: number, numInsert: number): void {
    const sheet = this.sheets[sheetName];
    if (sheet) {
      const inserts: null[] = [];
      for (let i = 0; i < numInsert; i++) {
        inserts.push(null);
      }
      _.each(sheet, (row: (CellData | undefined | null)[]) => {
        row.splice(start, numDelete, ...inserts);
      });
    }
  }
}

export default CellMatrix;
