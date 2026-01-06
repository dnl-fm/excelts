// =============================================================================
// The ZipWriter class
// Packs streamed data into an output zip stream using fflate
import { zipSync, strToU8 } from 'fflate';
import SimpleBuffer from './simple-buffer.ts';
import { isBytes } from './bytes.ts';

type EventHandler = (...args: unknown[]) => void;

class ZipWriter {
  options: Record<string, unknown>;
  files: Map<string, Uint8Array>;
  stream: SimpleBuffer;
  private handlers: Map<string, EventHandler[]> = new Map();

  constructor(options?: Record<string, unknown>) {
    this.options = options || {};
    this.files = new Map();
    this.stream = new SimpleBuffer();
  }

  on(event: string, handler: EventHandler): this {
    if (!this.handlers.has(event)) {
      this.handlers.set(event, []);
    }
    this.handlers.get(event)!.push(handler);
    return this;
  }

  emit(event: string, ...args: unknown[]): boolean {
    const handlers = this.handlers.get(event);
    if (handlers) {
      handlers.forEach(h => h(...args));
      return true;
    }
    return false;
  }

  append(data: string | Uint8Array, options: { name: string; base64?: boolean }): void {
    let bytes: Uint8Array;
    
    if (options.base64) {
      // Decode base64 string to bytes
      const binaryString = atob(data as string);
      bytes = new Uint8Array(binaryString.length);
      for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
      }
    } else if (typeof data === 'string') {
      bytes = strToU8(data);
    } else if (isBytes(data)) {
      bytes = data;
    } else {
      bytes = new Uint8Array(data as ArrayBuffer);
    }
    
    this.files.set(options.name, bytes);
  }

  async finalize(): Promise<void> {
    // Convert Map to object for fflate
    const filesObj: Record<string, Uint8Array> = {};
    for (const [name, data] of this.files) {
      filesObj[name] = data;
    }
    
    const zipped = zipSync(filesObj, { level: 6 });
    this.stream.end(new Uint8Array(zipped));
    this.emit('finish');
  }

  // ==========================================================================
  // Stream.Readable interface
  read(size?: number): Uint8Array | null {
    return this.stream.read(size);
  }

  setEncoding(encoding: string): void {
    this.stream.setEncoding(encoding);
  }

  pause(): void {
    this.stream.pause();
  }

  resume(): void {
    this.stream.resume();
  }

  isPaused(): boolean {
    return this.stream.isPaused();
  }

  pipe<T>(destination: T, _options?: { end?: boolean }): T {
    this.stream.pipe(destination as unknown as { write(chunk: Uint8Array | string, callback?: () => void): void; end(): void; });
    return destination;
  }

  unpipe(destination?: unknown): void {
    this.stream.unpipe(destination as { write(chunk: Uint8Array | string, callback?: () => void): void; end(): void; });
  }

  unshift(_chunk: Uint8Array): void {
    // Not implemented for SimpleBuffer
  }

  wrap(_stream: unknown): void {
    // Not implemented for SimpleBuffer
  }
}

// =============================================================================

export default {
  ZipWriter,
};
