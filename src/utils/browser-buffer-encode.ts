// eslint-disable-next-line node/no-unsupported-features/node-builtins
import { Buffer } from 'buffer';

const textEncoder = typeof TextEncoder === 'undefined' ? null : new TextEncoder('utf-8');

function stringToBuffer(str) {
  if (typeof str !== 'string') {
    return str;
  }
  if (textEncoder) {
    return Buffer.from(textEncoder.encode(str).buffer);
  }
  return Buffer.from(str);
}

export { stringToBuffer };