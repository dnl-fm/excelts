
import _ from '../../../utils/under-dash.ts';
import Range from '../../../doc/range.ts';
import colCache from '../../../utils/col-cache.ts';
import Enums from '../../../doc/enums.ts';

interface MergeRef {
  address: string;
  master: string;
}

interface RowWithCells {
  cells: Array<{
    type: number;
    address: string;
    master?: string;
  } | undefined>;
}

interface RangeLocation {
  top: number;
  left: number;
  bottom: number;
  right: number;
  tl: string;
  br: string;
  dimensions: string;
}

class Merges {
  merges: Record<string, Range>;
  hash: Record<string, Range>;

  constructor() {
    // optional mergeCells is array of ranges (like the xml)
    this.merges = {};
    this.hash = {};
  }

  add(merge: MergeRef): void {
    // merge is {address, master}
    if (this.merges[merge.master]) {
      this.merges[merge.master].expandToAddress(merge.address);
    } else {
      const range = `${merge.master}:${merge.address}`;
      this.merges[merge.master] = new Range(range);
    }
  }

  get mergeCells(): string[] {
    return _.map(this.merges, (merge: Range) => merge.range) as string[];
  }

  reconcile(mergeCells: string[], rows: RowWithCells[]): void {
    // reconcile merge list with merge cells
    _.each(mergeCells, (merge: string) => {
      const dimensions = colCache.decode(merge) as RangeLocation;
      for (let i = dimensions.top; i <= dimensions.bottom; i++) {
        const row = rows[i - 1];
        for (let j = dimensions.left; j <= dimensions.right; j++) {
          const cell = row.cells[j - 1];
          if (!cell) {
            // nulls are not included in document - so if master cell has no value - add a null one here
            row.cells[j] = {
              type: Enums.ValueType.Null,
              address: colCache.encodeAddress(i, j),
            };
          } else if (cell.type === Enums.ValueType.Merge) {
            cell.master = dimensions.tl;
          }
        }
      }
    });
  }

  getMasterAddress(address: string): string | undefined {
    // if address has been merged, return its master's address. Assumes reconcile has been called
    const range = this.hash[address];
    return range && range.tl;
  }
}

export default Merges;
