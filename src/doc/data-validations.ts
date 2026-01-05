/**
 * DataValidations stores per-cell validation rules for a worksheet.
 */
class DataValidations {
  model: Record<string, unknown>;

  constructor(model?: Record<string, unknown>) {
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