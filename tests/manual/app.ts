/* eslint-disable no-console */

import fs from 'fs';
import express from 'express';
import path from 'path';
import StreamBuf from '../../src/utils/stream-buf.ts';
import ExcelTS from '../../src/index.ts';

console.log('Copying bundle.js to public folder');
fs.createReadStream(`${__dirname}/../../dist/exceljs.min.js`).pipe(
  fs.createWriteStream(`${__dirname}/public/exceljs.min.js`)
);
fs.createReadStream(`${__dirname}/../../dist/exceljs.js`).pipe(
  fs.createWriteStream(`${__dirname}/public/exceljs.js`)
);

const app = express();

app.use('/', express.static(path.join(__dirname, 'public')));

app.post('/api/upload', (req, res) => {
  const wb = new ExcelTS.Workbook();

  const stream = new StreamBuf();
  stream.on('finish', () => {
    const base64 = stream.read();

    wb.xlsx.load(base64, {base64: true}).then(() => {
      const ws = wb.getWorksheet('blort');

      console.log('XLSX uploaded:');
      console.log('A1', ws.getCell('A1').value);
      console.log('A2', ws.getCell('A2').value);

      ws.getCell('A1').value = 'Hey Ho!';
      ws.getCell('A2').value = 14;

      const outStream = new StreamBuf();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (wb.xlsx.write as any)(outStream).then(() => {
        const b = outStream.read() as Buffer;
        const s = b.toString('base64');
        res.write(s);
        res.end();
      });
    });
  });

  req.pipe(stream);
});

app.listen(3003);
console.log('Listening on port 3003');
