import {expect, describe, it} from 'bun:test';

const WorksheetWriter = verquire('stream/xlsx/worksheet-writer');
const StreamBuf = verquire('utils/stream-buf');

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
          const xml = mockWorkbook.stream.read().toString();
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
