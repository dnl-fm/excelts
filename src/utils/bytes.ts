/**
 * Browser-native byte utilities to replace Node.js Buffer.
 */

const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder();

/**
 * Concatenate multiple Uint8Arrays into one.
 */
export function concat(arrays: Uint8Array[]): Uint8Array {
  if (arrays.length === 0) return new Uint8Array(0);
  if (arrays.length === 1) return arrays[0];

  const totalLength = arrays.reduce((sum, arr) => sum + arr.length, 0);
  const result = new Uint8Array(totalLength);
  let offset = 0;
  for (const arr of arrays) {
    result.set(arr, offset);
    offset += arr.length;
  }
  return result;
}

/**
 * Allocate a new Uint8Array of given size (zero-filled).
 */
export function alloc(size: number): Uint8Array {
  return new Uint8Array(size);
}

/**
 * Convert string to Uint8Array with specified encoding.
 */
export function fromString(str: string, encoding: 'utf8' | 'utf-8' | 'base64' | 'utf16le' | 'ucs2' = 'utf8'): Uint8Array {
  if (encoding === 'utf8' || encoding === 'utf-8') {
    return textEncoder.encode(str);
  }

  if (encoding === 'base64') {
    const binary = atob(str);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
  }

  if (encoding === 'utf16le' || encoding === 'ucs2') {
    const bytes = new Uint8Array(str.length * 2);
    for (let i = 0; i < str.length; i++) {
      const code = str.charCodeAt(i);
      bytes[i * 2] = code & 0xff;
      bytes[i * 2 + 1] = (code >> 8) & 0xff;
    }
    return bytes;
  }

  throw new Error(`Unsupported encoding: ${encoding}`);
}

/**
 * Convert Uint8Array to string with specified encoding.
 */
export function toString(bytes: Uint8Array, encoding: 'utf8' | 'utf-8' | 'base64' | 'utf16le' | 'ucs2' | 'hex' = 'utf8'): string {
  if (encoding === 'utf8' || encoding === 'utf-8') {
    return textDecoder.decode(bytes);
  }

  if (encoding === 'base64') {
    let binary = '';
    for (let i = 0; i < bytes.length; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
  }

  if (encoding === 'utf16le' || encoding === 'ucs2') {
    let result = '';
    for (let i = 0; i < bytes.length; i += 2) {
      result += String.fromCharCode(bytes[i] | (bytes[i + 1] << 8));
    }
    return result;
  }

  if (encoding === 'hex') {
    return Array.from(bytes)
      .map(b => b.toString(16).padStart(2, '0'))
      .join('');
  }

  throw new Error(`Unsupported encoding: ${encoding}`);
}

/**
 * Create Uint8Array from various inputs.
 */
export function toBytes(data: string | ArrayBuffer | Uint8Array | number[], encoding?: 'utf8' | 'utf-8' | 'base64' | 'utf16le' | 'ucs2'): Uint8Array {
  if (data instanceof Uint8Array) {
    return data;
  }
  if (data instanceof ArrayBuffer) {
    return new Uint8Array(data);
  }
  if (Array.isArray(data)) {
    return new Uint8Array(data);
  }
  if (typeof data === 'string') {
    return fromString(data, encoding || 'utf8');
  }
  throw new Error(`Cannot convert to bytes: ${typeof data}`);
}

/**
 * Check if value is a Uint8Array (replacement for Buffer.isBuffer).
 */
export function isBytes(value: unknown): value is Uint8Array {
  return value instanceof Uint8Array;
}
