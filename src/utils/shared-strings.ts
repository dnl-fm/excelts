class SharedStrings {
  private _values: unknown[];
  private _totalRefs: number;
  private _hash: Record<string, number>;

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
    let index = this._hash[String(value)];
    if (index === undefined) {
      index = this._hash[String(value)] = this._values.length;
      this._values.push(value);
    }
    this._totalRefs++;
    return index;
  }
}

export default SharedStrings;