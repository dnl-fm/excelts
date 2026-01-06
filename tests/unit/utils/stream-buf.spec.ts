import path from 'path';

import StreamBuf from '../../../src/utils/stream-buf.ts';
import StringBuf from '../../../src/utils/string-buf.ts';
import { toString } from '../../../src/utils/bytes.ts';

describe('StreamBuf', () => {
  // StreamBuf is designed as a general-purpose writable-readable stream
  // However its use in ExcelJS is primarily as a memory buffer between
  // the streaming writers and the archive, hence the tests here will
  // focus just on that.
  it('writes strings as UTF8', () => {
    const stream = new StreamBuf();
    stream.write('Hello, World!');
    const chunk = stream.read();
    expect(chunk instanceof Uint8Array).toBeTruthy();
    expect(toString(chunk, 'utf8')).toBe('Hello, World!');
  });

  it('writes StringBuf chunks', () => {
    const stream = new StreamBuf();
    const strBuf = new StringBuf({size: 64});
    strBuf.addText('Hello, World!');
    stream.write(strBuf);
    const chunk = stream.read();
    expect(chunk instanceof Uint8Array).toBeTruthy();
    expect(toString(chunk, 'utf8')).toBe('Hello, World!');
  });

  it('signals end', done => {
    const stream = new StreamBuf();
    stream.on('finish', () => {
      done();
    });
    stream.write('Hello, World!');
    stream.end();
  });

  it('handles buffers', async () => {
    const fileData = await Bun.file(path.join(__dirname, 'data/image1.png')).bytes();
    const sb = new StreamBuf();
    sb.write(fileData);
    sb.end();
    const buf = sb.toBuffer();
    expect(buf!.length).toBe(1672);
  });

  it('handle unsupported type of chunk', async () => {
    const stream = new StreamBuf();
    try {
      // @ts-expect-error Testing invalid input
      await stream.write({});
      throw new Error('should fail for given argument');
    } catch (e) {
      expect((e as Error).message).toBe(
        'Chunk must be one of type String, Uint8Array, ArrayBuffer or StringBuf.'
      );
    }
  });
});
