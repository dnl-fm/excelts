import {expect, describe, it} from 'bun:test';

import WorksheetWriter from '../../../src/stream/xlsx/worksheet-writer.ts';
import StreamBuf from '../../../src/utils/stream-buf.ts';
import { toString } from '../../../src/utils/bytes.ts';

/**
 * Basic XML validity check
 */
function expectXmlValid(xml: string): void {
  expect(xml.trim().startsWith('<')).toBe(true);
  expect(xml.trim().endsWith('>')).toBe(true);
}

describe('Workbook Writer', () => {
  it('generates valid xml even when there is no data', () =>
    new Promise<void>((resolve, reject) => {
      const mockWorkbook = {
        _openStream() {
          return this.stream;
        },
        stream: new StreamBuf(),
      };
      mockWorkbook.stream.on('finish', () => {
        try {
          const xml = toString(mockWorkbook.stream.read(), 'utf8');
          expectXmlValid(xml);
          resolve();
        } catch (error) {
          reject(error);
        }
      });

      const writer = new WorksheetWriter({
        id: 1,
        workbook: mockWorkbook,
      });

      writer.commit();
    }));
});
