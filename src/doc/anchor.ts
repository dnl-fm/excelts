
/**
 * Anchor defines an image or drawing position within a worksheet.
 */
import colCache from '../utils/col-cache.ts';

interface NativeAnchorAddress {
  nativeCol?: number;
  nativeColOff?: number;
  nativeRow?: number;
  nativeRowOff?: number;
}

interface ColRowAddress {
  col?: number;
  row?: number;
}

type AnchorAddress = string | NativeAnchorAddress | ColRowAddress;

interface WorksheetLike {
  getColumn(col: number): { width?: number; isCustomWidth?: boolean };
  getRow(row: number): { height?: number };
}

class Anchor {
  worksheet: WorksheetLike | undefined;
  nativeCol: number;
  nativeColOff: number;
  nativeRow: number;
  nativeRowOff: number;

  constructor(worksheet?: WorksheetLike, address?: AnchorAddress, offset = 0) {
    this.worksheet = worksheet;

    if (!address) {
      this.nativeCol = 0;
      this.nativeColOff = 0;
      this.nativeRow = 0;
      this.nativeRowOff = 0;
    } else if (typeof address === 'string') {
      const decoded = colCache.decodeAddress(address);
      this.nativeCol = decoded.col + offset;
      this.nativeColOff = 0;
      this.nativeRow = decoded.row + offset;
      this.nativeRowOff = 0;
    } else if ('nativeCol' in address && address.nativeCol !== undefined) {
      const native = address as NativeAnchorAddress;
      this.nativeCol = native.nativeCol || 0;
      this.nativeColOff = native.nativeColOff || 0;
      this.nativeRow = native.nativeRow || 0;
      this.nativeRowOff = native.nativeRowOff || 0;
    } else if ('col' in address && address.col !== undefined) {
      const colRow = address as ColRowAddress;
      this.col = (colRow.col || 0) + offset;
      this.row = (colRow.row || 0) + offset;
    } else {
      this.nativeCol = 0;
      this.nativeColOff = 0;
      this.nativeRow = 0;
      this.nativeRowOff = 0;
    }
  }

  static asInstance(model) {
    return model instanceof Anchor || model == null ? model : new Anchor(model);
  }

  get col() {
    return this.nativeCol + (Math.min(this.colWidth - 1, this.nativeColOff) / this.colWidth);
  }

  set col(v) {
    this.nativeCol = Math.floor(v);
    this.nativeColOff = Math.floor((v - this.nativeCol) * this.colWidth);
  }

  get row() {
    return this.nativeRow + (Math.min(this.rowHeight - 1, this.nativeRowOff) / this.rowHeight);
  }

  set row(v) {
    this.nativeRow = Math.floor(v);
    this.nativeRowOff = Math.floor((v - this.nativeRow) * this.rowHeight);
  }

  get colWidth() {
    return this.worksheet &&
      this.worksheet.getColumn(this.nativeCol + 1) &&
      this.worksheet.getColumn(this.nativeCol + 1).isCustomWidth
      ? Math.floor(this.worksheet.getColumn(this.nativeCol + 1).width * 10000)
      : 640000;
  }

  get rowHeight() {
    return this.worksheet &&
      this.worksheet.getRow(this.nativeRow + 1) &&
      this.worksheet.getRow(this.nativeRow + 1).height
      ? Math.floor(this.worksheet.getRow(this.nativeRow + 1).height * 10000)
      : 180000;
  }

  get model() {
    return {
      nativeCol: this.nativeCol,
      nativeColOff: this.nativeColOff,
      nativeRow: this.nativeRow,
      nativeRowOff: this.nativeRowOff,
    };
  }

  set model(value) {
    this.nativeCol = value.nativeCol;
    this.nativeColOff = value.nativeColOff;
    this.nativeRow = value.nativeRow;
    this.nativeRowOff = value.nativeRowOff;
  }
}

export default Anchor;