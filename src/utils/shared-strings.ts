class SharedStrings {
  constructor() {
    this._values = [];
    this._totalRefs = 0;
    this._hash = Object.create(null);
  }

  get count(): number {
    return this._values.length;
  }

  get values(): unknown[] {
    return this._values;
  }

  get totalRefs(): number {
    return this._totalRefs;
  }

  getString(index: number): unknown {
    return this._values[index];
  }

  add(value: unknown): number {
    let index = this._hash[value];
    if (index === undefined) {
      index = this._hash[value] = this._values.length;
      this._values.push(value);
    }
    this._totalRefs++;
    return index;
  }
}

export default SharedStrings;