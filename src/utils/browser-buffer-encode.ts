import { toBytes } from './bytes.ts';

function stringToBuffer(str: string | Uint8Array): Uint8Array {
  if (typeof str !== 'string') {
    return str;
  }
  return toBytes(str, 'utf8');
}

export { stringToBuffer };
