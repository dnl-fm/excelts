
/**
 * ModelContainer wraps a workbook-like model for XLSX IO.
 */
import XLSX from '../xlsx/xlsx.ts';

class ModelContainer {
  constructor(model) {
    this.model = model;
  }

  get xlsx() {
    if (!this._xlsx) {
      this._xlsx = new XLSX(this);
    }
    return this._xlsx;
  }
}

export default ModelContainer;