import { concat, fromString, toString, alloc } from './bytes.ts';

// Browser-compatible crypto using SubtleCrypto
const Encryptor = {
  /**
   * Calculate a hash of the concatenated buffers with the given algorithm.
   * @param algorithm - The hash algorithm (e.g., 'SHA-256', 'SHA-512')
   * @param buffers - The buffers to hash
   * @returns The hash as Uint8Array
   */
  async hash(algorithm: string, ...buffers: Uint8Array[]): Promise<Uint8Array> {
    // Normalize algorithm name for SubtleCrypto
    const algoMap: Record<string, string> = {
      'sha1': 'SHA-1',
      'sha256': 'SHA-256',
      'sha384': 'SHA-384',
      'sha512': 'SHA-512',
    };
    const normalizedAlgo = algoMap[algorithm.toLowerCase()] || algorithm.toUpperCase();
    
    const combined = concat(buffers);
    const hashBuffer = await crypto.subtle.digest(normalizedAlgo, combined as BufferSource);
    return new Uint8Array(hashBuffer);
  },

  /**
   * Synchronous hash using SubtleCrypto (for compatibility)
   * Note: This is async internally but wrapped for sync interface
   */
  hashSync(_algorithm: string, ..._buffers: Uint8Array[]): Uint8Array {
    // For browser compatibility, we need async crypto
    // This is a placeholder that will be called in async context
    throw new Error('Use async hash() method instead of hashSync()');
  },

  /**
   * Convert a password into an encryption key
   * @param password - The password
   * @param hashAlgorithm - The hash algorithm
   * @param saltValue - The salt value (base64 encoded)
   * @param spinCount - The spin count
   * @returns The encryption key as base64 string
   */
  async convertPasswordToHash(
    password: string,
    hashAlgorithm: string,
    saltValue: string,
    spinCount: number
  ): Promise<string> {
    // Password must be in unicode buffer (UTF-16LE)
    const passwordBuffer = fromString(password, 'utf16le');
    
    // Decode salt from base64
    const saltBuffer = fromString(saltValue, 'base64');
    
    // Generate the initial hash
    let key = await this.hash(hashAlgorithm, saltBuffer, passwordBuffer);
    
    // Now regenerate until spin count
    for (let i = 0; i < spinCount; i++) {
      const iterator = alloc(4);
      // Write little-endian uint32
      iterator[0] = i & 0xff;
      iterator[1] = (i >> 8) & 0xff;
      iterator[2] = (i >> 16) & 0xff;
      iterator[3] = (i >> 24) & 0xff;
      key = await this.hash(hashAlgorithm, key, iterator);
    }
    
    return toString(key, 'base64');
  },

  /**
   * Generates cryptographically strong pseudo-random data.
   * @param size The size argument is a number indicating the number of bytes to generate.
   */
  randomBytes(size: number): Uint8Array {
    const bytes = new Uint8Array(size);
    crypto.getRandomValues(bytes);
    return bytes;
  },
};

export default Encryptor;
