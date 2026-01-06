// Browser tests - using Bun environment instead of Jasmine
import ExcelTS from '../../src/index.ts';
import { toString as bytesToString, fromString } from '../../src/utils/bytes.ts';

function unexpectedError(done) {
  return function(error) {
    // eslint-disable-next-line no-console
    console.error('Error Caught', error.message, error.stack);
    expect(true).toEqual(false);
    done();
  };
}

describe('ExcelTS', () => {
  it('should read and write xlsx via binary buffer', done => {
    const wb = new ExcelTS.Workbook();
    const ws = wb.addWorksheet('blort');

    ws.getCell('A1').value = 'Hello, World!';
    ws.getCell('A2').value = 7;

    wb.xlsx
      .writeBuffer()
      .then(buffer => {
        const wb2 = new ExcelTS.Workbook();
        return wb2.xlsx.load(buffer).then(() => {
          const ws2 = wb2.getWorksheet('blort');
          expect(ws2).toBeTruthy();

          expect(ws2.getCell('A1').value).toEqual('Hello, World!');
          expect(ws2.getCell('A2').value).toEqual(7);
          done();
        });
      })
      .catch(error => {
        throw error;
      })
      .catch(unexpectedError(done));
  });
  it('should read and write xlsx via base64 buffer', done => {
    const options = {
      base64: true,
    };
    const wb = new ExcelTS.Workbook();
    const ws = wb.addWorksheet('blort');

    ws.getCell('A1').value = 'Hello, World!';
    ws.getCell('A2').value = 7;

    wb.xlsx
      .writeBuffer()
      .then(buffer => {
        // Convert Uint8Array to base64 string
        const base64String = bytesToString(buffer, 'base64');
        // Convert base64 string back to Uint8Array for load
        const base64Buffer = fromString(base64String, 'utf8');
        
        const wb2 = new ExcelTS.Workbook();
        return wb2.xlsx.load(base64Buffer, options).then(() => {
          const ws2 = wb2.getWorksheet('blort');
          expect(ws2).toBeTruthy();

          expect(ws2.getCell('A1').value).toEqual('Hello, World!');
          expect(ws2.getCell('A2').value).toEqual(7);
          done();
        });
      })
      .catch(error => {
        throw error;
      })
      .catch(unexpectedError(done));
  });
  it('should write csv via buffer', done => {
    const wb = new ExcelTS.Workbook();
    const ws = wb.addWorksheet('blort');

    ws.getCell('A1').value = 'Hello, World!';
    ws.getCell('B1').value = 'What time is it?';
    ws.getCell('A2').value = 7;
    ws.getCell('B2').value = '12pm';

    wb.csv
      .writeBuffer()
      .then(buffer => {
        expect(bytesToString(buffer, 'utf8').trim()).toEqual(
          '"Hello, World!",What time is it?\n7,12pm'
        );
        done();
      })
      .catch(error => {
        throw error;
      })
      .catch(unexpectedError(done));
  });
});
