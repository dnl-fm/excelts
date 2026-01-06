import {PassThrough} from 'stream';
import {expect, describe, it} from 'bun:test';
import ExcelTS from '../../src/index.ts';

describe('typescript', () => {
  it('can create and buffer xlsx', async () => {
    const wb = new ExcelTS.Workbook();
    const ws = wb.addWorksheet('blort');
    ws.getCell('A1').value = 7;
    const buffer = await wb.xlsx.writeBuffer({
      useStyles: true,
      useSharedStrings: true,
    });

    const wb2 = new ExcelTS.Workbook();
    await wb2.xlsx.load(buffer);
    const ws2 = wb2.getWorksheet('blort');
    expect(ws2.getCell('A1').value).toBe(7);
  });

  it('can create and stream xlsx', async () => {
    const wb = new ExcelTS.Workbook();
    const ws = wb.addWorksheet('blort');
    ws.getCell('A1').value = 7;

    const wb2 = new ExcelTS.Workbook();
    const stream = new PassThrough();
    const readPromise = wb2.xlsx.read(stream);

    await wb.xlsx.write(stream);
    stream.end();

    await readPromise;
    const ws2 = wb2.getWorksheet('blort');
    expect(ws2.getCell('A1').value).toBe(7);
  });
});
