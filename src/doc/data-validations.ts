/**
 * DataValidations stores per-cell validation rules for a worksheet.
 */
class DataValidations {
  constructor(model) {
    this.model = model || {};
  }

  add(address, validation) {
    return (this.model[address] = validation);
  }

  find(address) {
    return this.model[address];
  }

  remove(address) {
    this.model[address] = undefined;
  }
}

export default DataValidations;