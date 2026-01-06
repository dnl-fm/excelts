// StringBuilder - a way to keep string memory operations to a minimum
// while building the strings for the xml files
class StringBuilder {
  _buf: string[];

  constructor() {
    this._buf = [];
    this.reset();
  }

  get length() {
    return this._buf.length;
  }

  toString() {
    return this._buf.join('');
  }

  reset(position?: number) {
    if (position) {
      while (this._buf.length > position) {
        this._buf.pop();
      }
    } else {
      this._buf = [];
    }
  }

  addText(text) {
    this._buf.push(text);
  }

  addStringBuf(inBuf) {
    this._buf.push(inBuf.toString());
  }
}

export default StringBuilder;