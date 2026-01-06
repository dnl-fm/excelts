/**
 * StreamBuf - Browser-native multi-purpose read-write buffer
 * Replaces Node.js Stream.Duplex with Uint8Array-based implementation
 */

import { concat, alloc, isBytes, fromString } from './bytes.ts';
import SimpleEventEmitter from './event-emitter.ts';
import StringBuf from './string-buf.ts';

// =============================================================================
// Chunk classes - encapsulate incoming data

class StringChunk {
  private _data: string;
  private _encoding: 'utf8' | 'base64' | 'utf16le';
  private _buffer: Uint8Array | undefined;

  constructor(data: string, encoding?: string) {
    this._data = data;
    this._encoding = (encoding as 'utf8' | 'base64' | 'utf16le') || 'utf8';
  }

  get length(): number {
    return this.toBuffer().length;
  }

  copy(target: Uint8Array, targetOffset: number, offset: number, length: number): number {
    const buf = this.toBuffer();
    const toCopy = buf.subarray(offset, offset + length);
    target.set(toCopy, targetOffset);
    return toCopy.length;
  }

  toBuffer(): Uint8Array {
    if (!this._buffer) {
      this._buffer = fromString(this._data, this._encoding);
    }
    return this._buffer;
  }
}

class StringBufChunk {
  private _data: StringBuf;

  constructor(data: StringBuf) {
    this._data = data;
  }

  get length(): number {
    return this._data.length;
  }

  copy(target: Uint8Array, targetOffset: number, offset: number, length: number): number {
    const buf = this._data._buf;
    const toCopy = buf.subarray(offset, offset + length);
    target.set(toCopy, targetOffset);
    return toCopy.length;
  }

  toBuffer(): Uint8Array {
    return this._data.toBuffer();
  }
}

class BufferChunk {
  private _data: Uint8Array;

  constructor(data: Uint8Array) {
    this._data = data;
  }

  get length(): number {
    return this._data.length;
  }

  copy(target: Uint8Array, targetOffset: number, offset: number, length: number): number {
    const toCopy = this._data.subarray(offset, offset + length);
    target.set(toCopy, targetOffset);
    return toCopy.length;
  }

  toBuffer(): Uint8Array {
    return this._data;
  }
}

type Chunk = StringChunk | StringBufChunk | BufferChunk;

// =============================================================================
// ReadWriteBuf - a single buffer supporting simple read-write

class ReadWriteBuf {
  size: number;
  buffer: Uint8Array;
  iRead: number;
  iWrite: number;

  constructor(size: number) {
    this.size = size;
    this.buffer = alloc(size);
    this.iRead = 0;
    this.iWrite = 0;
  }

  toBuffer(): Uint8Array {
    if (this.iRead === 0 && this.iWrite === this.size) {
      return this.buffer;
    }
    return this.buffer.slice(this.iRead, this.iWrite);
  }

  get length(): number {
    return this.iWrite - this.iRead;
  }

  get eod(): boolean {
    return this.iRead === this.iWrite;
  }

  get full(): boolean {
    return this.iWrite === this.size;
  }

  read(size?: number): Uint8Array | null {
    if (size === 0) {
      return null;
    }

    if (size === undefined || size >= this.length) {
      const buf = this.toBuffer();
      this.iRead = this.iWrite;
      return buf;
    }

    const buf = this.buffer.slice(this.iRead, this.iRead + size);
    this.iRead += size;
    return buf;
  }

  write(chunk: Chunk, offset: number, length: number): number {
    const size = Math.min(length, this.size - this.iWrite);
    chunk.copy(this.buffer, this.iWrite, offset, size);
    this.iWrite += size;
    return size;
  }
}

// =============================================================================
// PipeDestination interface

interface PipeDestination {
  write(chunk: Uint8Array, callback?: () => void): boolean;
  end(): void;
}

// =============================================================================
// StreamBuf - a multi-purpose read-write stream
//  As MemBuf - write as much data as you like. Then call toBuffer() to consolidate
//  As StreamHub - pipe to multiple writables
//  As readable stream - feed data into the writable part and have some other code read from it.

interface StreamBufOptions {
  bufSize?: number;
  batch?: boolean;
}

class StreamBuf extends SimpleEventEmitter {
  bufSize: number;
  buffers: ReadWriteBuf[];
  batch: boolean;
  corked: boolean;
  inPos: number;
  outPos: number;
  pipes: PipeDestination[];
  paused: boolean;
  encoding: string | null;

  constructor(options?: StreamBufOptions) {
    super();
    options = options || {};
    this.bufSize = options.bufSize || 1024 * 1024;
    this.buffers = [];
    this.batch = options.batch || false;
    this.corked = false;
    this.inPos = 0;
    this.outPos = 0;
    this.pipes = [];
    this.paused = false;
    this.encoding = null;
  }

  toBuffer(): Uint8Array | null {
    switch (this.buffers.length) {
      case 0:
        return null;
      case 1:
        return this.buffers[0].toBuffer();
      default:
        return concat(this.buffers.map(rwBuf => rwBuf.toBuffer()));
    }
  }

  private _getWritableBuffer(): ReadWriteBuf {
    if (this.buffers.length) {
      const last = this.buffers[this.buffers.length - 1];
      if (!last.full) {
        return last;
      }
    }
    const buf = new ReadWriteBuf(this.bufSize);
    this.buffers.push(buf);
    return buf;
  }

  private async _pipe(chunk: Chunk): Promise<void> {
    const writePromises = this.pipes.map(pipe => {
      return new Promise<void>(resolve => {
        pipe.write(chunk.toBuffer(), () => {
          resolve();
        });
      });
    });
    await Promise.all(writePromises);
  }

  private _writeToBuffers(chunk: Chunk): void {
    let inPos = 0;
    const inLen = chunk.length;
    while (inPos < inLen) {
      const buffer = this._getWritableBuffer();
      inPos += buffer.write(chunk, inPos, inLen - inPos);
    }
  }

  async write(
    data: string | Uint8Array | ArrayBuffer | StringBuf,
    encoding?: string | (() => void),
    callback?: () => void
  ): Promise<boolean> {
    if (typeof encoding === 'function') {
      callback = encoding;
      encoding = 'utf8';
    }
    callback = callback || (() => {});

    // encapsulate data into a chunk
    let chunk: Chunk;
    if (data instanceof StringBuf) {
      chunk = new StringBufChunk(data);
    } else if (isBytes(data)) {
      chunk = new BufferChunk(data);
    } else if (data instanceof ArrayBuffer) {
      chunk = new BufferChunk(new Uint8Array(data));
    } else if (typeof data === 'string') {
      chunk = new StringChunk(data, encoding as string);
    } else {
      throw new Error('Chunk must be one of type String, Uint8Array, ArrayBuffer or StringBuf.');
    }

    // now, do something with the chunk
    if (this.pipes.length) {
      if (this.batch) {
        this._writeToBuffers(chunk);
        while (!this.corked && this.buffers.length > 1) {
          await this._pipe(new BufferChunk(this.buffers.shift()!.toBuffer()));
        }
      } else if (!this.corked) {
        await this._pipe(chunk);
        callback();
      } else {
        this._writeToBuffers(chunk);
        queueMicrotask(callback);
      }
    } else {
      if (!this.paused) {
        this.emit('data', chunk.toBuffer());
      }

      this._writeToBuffers(chunk);
      this.emit('readable');
    }

    return true;
  }

  cork(): void {
    this.corked = true;
  }

  private _flush(): void {
    if (this.pipes.length) {
      while (this.buffers.length) {
        const buf = this.buffers.shift()!;
        this._pipe(new BufferChunk(buf.toBuffer()));
      }
    }
  }

  uncork(): void {
    this.corked = false;
    this._flush();
  }

  end(
    chunk?: string | Uint8Array | ArrayBuffer | StringBuf,
    encoding?: string | (() => void),
    callback?: () => void
  ): void {
    if (typeof encoding === 'function') {
      callback = encoding;
      encoding = undefined;
    }

    const writeComplete = (error?: Error) => {
      if (error && callback) {
        callback();
      } else {
        this._flush();
        this.pipes.forEach(pipe => {
          pipe.end();
        });
        this.emit('finish');
        if (callback) callback();
      }
    };

    if (chunk) {
      this.write(chunk, encoding, writeComplete);
    } else {
      writeComplete();
    }
  }

  read(size?: number): Uint8Array {
    if (size) {
      const buffers: Uint8Array[] = [];
      let remaining = size;
      while (remaining && this.buffers.length && !this.buffers[0].eod) {
        const first = this.buffers[0];
        const buffer = first.read(remaining);
        if (buffer) {
          remaining -= buffer.length;
          buffers.push(buffer);
        }
        if (first.eod && first.full) {
          this.buffers.shift();
        }
      }
      return concat(buffers);
    }

    const buffers = this.buffers.map(buf => buf.toBuffer()).filter(Boolean) as Uint8Array[];
    this.buffers = [];
    return concat(buffers);
  }

  setEncoding(encoding: string): void {
    this.encoding = encoding;
  }

  pause(): void {
    this.paused = true;
  }

  resume(): void {
    this.paused = false;
  }

  isPaused(): boolean {
    return this.paused;
  }

  pipe(destination: PipeDestination): PipeDestination {
    this.pipes.push(destination);
    if (!this.paused && this.buffers.length) {
      this.end();
    }
    return destination;
  }

  unpipe(destination?: PipeDestination): void {
    if (destination) {
      this.pipes = this.pipes.filter(pipe => pipe !== destination);
    } else {
      this.pipes = [];
    }
  }

  unshift(_chunk: unknown): void {
    throw new Error('Not Implemented');
  }

  wrap(_stream: unknown): void {
    throw new Error('Not Implemented');
  }
}

export default StreamBuf;
