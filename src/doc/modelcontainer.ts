/**
 * ModelContainer wraps a workbook-like model for XLSX IO.
 */
import XLSX from '../xlsx/xlsx.ts';
import type Workbook from './workbook.ts';

class ModelContainer {
  model: unknown;
  private _xlsx?: XLSX;

  constructor(model: unknown) {
    this.model = model;
  }

  get xlsx(): XLSX {
    if (!this._xlsx) {
      // Cast to Workbook since XLSX only accesses model property which we have
      this._xlsx = new XLSX(this as unknown as Workbook);
    }
    return this._xlsx;
  }
}

export default ModelContainer;
