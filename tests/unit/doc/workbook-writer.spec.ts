
import Stream from 'readable-stream';

import Excel from '../../../src/index.ts';
describe('Workbook Writer', () => {
  it('returns undefined for non-existant sheet', () => {
    const stream = new Stream.Writable({
      write: function noop() {},
    });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wb = new Excel.stream.xlsx.WorkbookWriter({stream} as any);
    wb.addWorksheet('first');
    expect(wb.getWorksheet('w00t')).toBe(undefined);
  });
});
