import {describe, it, beforeAll, afterAll} from 'bun:test';
import {PassThrough} from 'stream';
import express from 'express';
import testutils from '../utils/index.ts';

import Excel from '../../src/index.ts';
describe('Express', () => {
  let server: ReturnType<typeof express.application.listen>;

  beforeAll(() => {
    const app = express();
    app.get('/workbook', (req, res) => {
      const wb = testutils.createTestBook(new Excel.Workbook(), 'xlsx');
      res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      );
      res.setHeader('Content-Disposition', 'attachment; filename=Report.xlsx');
      wb.xlsx.write(res).then(() => {
        res.end();
      });
    });
    server = app.listen(3003);
  });

  afterAll(() => {
    server.close();
  });

  it('downloads a workbook', async () => {
    const response = await fetch('http://127.0.0.1:3003/workbook');
    const arrayBuffer = await response.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    
    const wb2 = new Excel.Workbook();
    const stream = new PassThrough();
    stream.end(buffer);
    
    await wb2.xlsx.read(stream);
    testutils.checkTestBook(wb2, 'xlsx');
  });
});
