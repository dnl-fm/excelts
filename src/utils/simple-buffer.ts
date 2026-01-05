/**
 * SimpleBuffer - A Bun-native buffer accumulator
 * Replaces StreamBuf for simple write/read use cases without Node streams
 */

type PipeDestination = {
  write(chunk: Buffer | string, callback?: () => void): void;
  end(): void;
};

class SimpleBuffer {
  private chunks: Buffer[] = [];
  private pipes: PipeDestination[] = [];
  private finished = false;

  /**
   * Write data to the buffer
   */
  write(data: Buffer | string | ArrayBuffer, encoding?: string, callback?: () => void): boolean {
    if (typeof encoding === 'function') {
      callback = encoding;
      encoding = undefined;
    }

    let chunk: Buffer;
    if (data instanceof Buffer) {
      chunk = data;
    } else if (data instanceof ArrayBuffer) {
      chunk = Buffer.from(data);
    } else if (typeof data === 'string') {
      chunk = Buffer.from(data, (encoding as BufferEncoding) || 'utf8');
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
  end(data?: Buffer | string, encoding?: string, callback?: () => void): void {
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
   * Read accumulated data as Buffer
   */
  read(): Buffer {
    if (this.chunks.length === 0) return Buffer.alloc(0);
    if (this.chunks.length === 1) return this.chunks[0];
    return Buffer.concat(this.chunks);
  }

  /**
   * Get accumulated data as Buffer (alias for read)
   */
  toBuffer(): Buffer {
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
