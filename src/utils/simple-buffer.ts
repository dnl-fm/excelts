/**
 * SimpleBuffer - A browser-native buffer accumulator
 * Replaces StreamBuf for simple write/read use cases without Node streams
 */

import { concat, alloc, toBytes, isBytes } from './bytes.ts';

type PipeDestination = {
  write(chunk: Uint8Array | string, callback?: () => void): void;
  end(): void;
};

class SimpleBuffer {
  private chunks: Uint8Array[] = [];
  private pipes: PipeDestination[] = [];
  private finished = false;

  /**
   * Write data to the buffer
   */
  write(data: Uint8Array | string | ArrayBuffer, encoding?: string, callback?: () => void): boolean {
    if (typeof encoding === 'function') {
      callback = encoding;
      encoding = undefined;
    }

    let chunk: Uint8Array;
    if (isBytes(data)) {
      chunk = data;
    } else if (data instanceof ArrayBuffer) {
      chunk = new Uint8Array(data);
    } else if (typeof data === 'string') {
      chunk = toBytes(data, (encoding as 'utf8' | 'base64' | 'utf16le') || 'utf8');
    } else {
      throw new Error('Unsupported data type');
    }

    this.chunks.push(chunk);

    // Forward to pipes
    for (const dest of this.pipes) {
      dest.write(chunk);
    }

    if (callback) callback();
    return true;
  }

  /**
   * Signal end of writing
   */
  end(data?: Uint8Array | string, encoding?: string, callback?: () => void): void {
    if (data) {
      this.write(data, encoding, callback);
    }
    this.finished = true;
    
    // End all pipes
    for (const dest of this.pipes) {
      dest.end();
    }
  }

  /**
   * Read accumulated data as Uint8Array
   * @param _size - ignored, reads all data
   */
  read(_size?: number): Uint8Array {
    if (this.chunks.length === 0) return alloc(0);
    if (this.chunks.length === 1) return this.chunks[0];
    return concat(this.chunks);
  }

  /**
   * Get accumulated data as Uint8Array (alias for read)
   */
  toBuffer(): Uint8Array {
    return this.read();
  }

  /**
   * Pipe to a destination
   */
  pipe(destination: PipeDestination): PipeDestination {
    this.pipes.push(destination);
    
    // Write existing data to the new pipe
    for (const chunk of this.chunks) {
      destination.write(chunk);
    }
    
    if (this.finished) {
      destination.end();
    }
    
    return destination;
  }

  /**
   * Remove a pipe destination
   */
  unpipe(destination?: PipeDestination): void {
    if (destination) {
      this.pipes = this.pipes.filter(p => p !== destination);
    } else {
      this.pipes = [];
    }
  }

  // Compatibility methods (no-ops for interface compatibility)
  setEncoding(_encoding: string): void {}
  pause(): void {}
  resume(): void {}
  isPaused(): boolean { return false; }
  cork(): void {}
  uncork(): void {}
  on(_event: string, _handler: (...args: unknown[]) => void): this { return this; }
  emit(_event: string, ..._args: unknown[]): boolean { return true; }
}

export default SimpleBuffer;
