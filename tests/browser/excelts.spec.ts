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

// Mock document for download() tests
function createMockDocument() {
  const clicks: string[] = [];
  const appendedElements: any[] = [];
  const removedElements: any[] = [];
  
  return {
    clicks,
    appendedElements,
    removedElements,
    createElement: (tag: string) => ({
      tag,
      href: '',
      download: '',
      click: function() { clicks.push(this.download); },
    }),
    body: {
      appendChild: (el: any) => appendedElements.push(el),
      removeChild: (el: any) => removedElements.push(el),
    },
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

  it('should write xlsx as Blob with correct MIME type', async () => {
    const wb = new ExcelTS.Workbook();
    const ws = wb.addWorksheet('test');
    ws.getCell('A1').value = 'Hello';

    const blob = await wb.xlsx.writeBlob();

    expect(blob).toBeInstanceOf(Blob);
    expect(blob.type).toEqual('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    expect(blob.size).toBeGreaterThan(0);
  });

  it('should write csv as Blob with correct MIME type', async () => {
    const wb = new ExcelTS.Workbook();
    const ws = wb.addWorksheet('test');
    ws.getCell('A1').value = 'Hello';
    ws.getCell('B1').value = 'World';

    const blob = await wb.csv.writeBlob();

    expect(blob).toBeInstanceOf(Blob);
    expect(blob.type).toEqual('text/csv;charset=utf-8');
    expect(blob.size).toBeGreaterThan(0);
    
    // Verify content
    const text = await blob.text();
    expect(text.trim()).toEqual('Hello,World');
  });

  it('should download xlsx with correct filename', async () => {
    const mockDoc = createMockDocument();
    const revokedUrls: string[] = [];
    
    // Inject mocks
    (globalThis as any).document = mockDoc;
    const originalRevoke = URL.revokeObjectURL;
    URL.revokeObjectURL = (url: string) => revokedUrls.push(url);

    try {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('test');
      ws.getCell('A1').value = 'Hello';

      await wb.xlsx.download('report.xlsx');

      expect(mockDoc.clicks).toEqual(['report.xlsx']);
      expect(mockDoc.appendedElements.length).toEqual(1);
      expect(mockDoc.removedElements.length).toEqual(1);
      expect(revokedUrls.length).toEqual(1);
    } finally {
      delete (globalThis as any).document;
      URL.revokeObjectURL = originalRevoke;
    }
  });

  it('should append .xlsx extension if missing', async () => {
    const mockDoc = createMockDocument();
    
    (globalThis as any).document = mockDoc;
    const originalRevoke = URL.revokeObjectURL;
    URL.revokeObjectURL = () => {};

    try {
      const wb = new ExcelTS.Workbook();
      wb.addWorksheet('test');

      await wb.xlsx.download('report');

      expect(mockDoc.clicks).toEqual(['report.xlsx']);
    } finally {
      delete (globalThis as any).document;
      URL.revokeObjectURL = originalRevoke;
    }
  });

  it('should download csv with correct filename', async () => {
    const mockDoc = createMockDocument();
    
    (globalThis as any).document = mockDoc;
    const originalRevoke = URL.revokeObjectURL;
    URL.revokeObjectURL = () => {};

    try {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('test');
      ws.getCell('A1').value = 'Data';

      await wb.csv.download('export.csv');

      expect(mockDoc.clicks).toEqual(['export.csv']);
      expect(mockDoc.appendedElements.length).toEqual(1);
    } finally {
      delete (globalThis as any).document;
      URL.revokeObjectURL = originalRevoke;
    }
  });

  it('should append .csv extension if missing', async () => {
    const mockDoc = createMockDocument();
    
    (globalThis as any).document = mockDoc;
    const originalRevoke = URL.revokeObjectURL;
    URL.revokeObjectURL = () => {};

    try {
      const wb = new ExcelTS.Workbook();
      wb.addWorksheet('test');

      await wb.csv.download('data');

      expect(mockDoc.clicks).toEqual(['data.csv']);
    } finally {
      delete (globalThis as any).document;
      URL.revokeObjectURL = originalRevoke;
    }
  });
});
