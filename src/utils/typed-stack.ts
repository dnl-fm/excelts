class TypedStack {
  private _type: any;
  private _stack: unknown[];

  constructor(type: any) {
    this._type = type;
    this._stack = [];
  }

  get size(): number {
    return this._stack.length;
  }

  pop(): unknown {
    const tos = this._stack.pop();
    return tos || new this._type();
  }

  push(instance: unknown): void {
    if (!(instance instanceof this._type)) {
      throw new Error('Invalid type pushed to TypedStack');
    }
    this._stack.push(instance);
  }
}

export default TypedStack;