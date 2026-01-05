/**
 * ModelContainer wraps a workbook-like model for XLSX IO.
 */
import XLSX from '../xlsx/xlsx.ts';

class ModelContainer {
  model: unknown;
  private _xlsx?: XLSX;

  constructor(model: unknown) {
    this.model = model;
  }

  get xlsx(): XLSX {
    if (!this._xlsx) {
      this._xlsx = new XLSX(this);
    }
    return this._xlsx;
  }
}

export default ModelContainer;
