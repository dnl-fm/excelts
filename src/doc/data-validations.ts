/**
 * DataValidations stores per-cell validation rules for a worksheet.
 */
class DataValidations {
  constructor(model?: any) {
    this.model = model || {};
  }

  add(address: string, validation: unknown): unknown {
    return (this.model[address] = validation);
  }

  find(address: string): unknown {
    return this.model[address];
  }

  remove(address: string): void {
    this.model[address] = undefined;
  }
}

export default DataValidations;