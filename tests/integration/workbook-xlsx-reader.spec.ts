

import fs from 'fs';
import testutils from '../utils/index.ts';

import ExcelTS from '../../src/index.ts';
const TEST_FILE_NAME = './tests/out/wb.test.xlsx';

// need some architectural changes to make stream read work properly
// because of: shared strings, sheet names, etc are not read in guaranteed order
describe('WorkbookReader', () => {
  describe('Serialise', () => {
    it('xlsx file', () => {
      
      const wb = testutils.createTestBook(new ExcelTS.Workbook(), 'xlsx');

      return wb.xlsx
        .writeFile(TEST_FILE_NAME)
        .then(() => testutils.checkTestBookReader(TEST_FILE_NAME));
    });
  });

  describe('#readFile', () => {
    describe('Row limit', () => {
      it('should bail out if the file contains more rows than the limit', () => {
        const workbook = new ExcelTS.Workbook();
        // The Fibonacci sheet has 19 rows
        return workbook.xlsx
          .readFile('./tests/integration/data/fibonacci.xlsx', {maxRows: 10})
          .then(
            () => {
              throw new Error('Promise unexpectedly fulfilled');
            },
            err => {
              expect(err.message).toBe('Max row count (10) exceeded');
            }
          );
      });

      it('should fail fast on a huge file', () => {
        
        const workbook = new ExcelTS.Workbook();
        return workbook.xlsx
          .readFile('./tests/integration/data/huge.xlsx', {maxRows: 100})
          .then(
            () => {
              throw new Error('Promise unexpectedly fulfilled');
            },
            err => {
              expect(err.message).toBe('Max row count (100) exceeded');
            }
          );
      });

      it('should parse fine if the limit is not exceeded', () => {
        const workbook = new ExcelTS.Workbook();
        return workbook.xlsx.readFile(
          './tests/integration/data/fibonacci.xlsx',
          {maxRows: 20}
        );
      });
    });

    describe('Column limit', () => {
      it('should bail out if the file contains more cells than the limit', () => {
        const workbook = new ExcelTS.Workbook();
        // The many-columns sheet has 20 columns in row 2
        return workbook.xlsx
          .readFile('./tests/integration/data/many-columns.xlsx', {
            maxCols: 15,
          })
          .then(
            () => {
              throw new Error('Promise unexpectedly fulfilled');
            },
            err => {
              expect(err.message).toBe('Max column count (15) exceeded');
            }
          );
      });

      it('should fail fast on a huge file', () => {
        
        const workbook = new ExcelTS.Workbook();
        return workbook.xlsx
          .readFile('./tests/integration/data/huge.xlsx', {maxCols: 10})
          .then(
            () => {
              throw new Error('Promise unexpectedly fulfilled');
            },
            err => {
              expect(err.message).toBe('Max column count (10) exceeded');
            }
          );
      });

      it('should parse fine if the limit is not exceeded', () => {
        const workbook = new ExcelTS.Workbook();
        return workbook.xlsx.readFile(
          './tests/integration/data/many-columns.xlsx',
          {maxCols: 40}
        );
      });
    });
  });

  describe('#read', () => {
    describe('Row limit', () => {
      it('should bail out if the file contains more rows than the limit', () => {
        const workbook = new ExcelTS.Workbook();
        // The Fibonacci sheet has 19 rows
        return workbook.xlsx
          .read(fs.createReadStream('./tests/integration/data/fibonacci.xlsx'), {
            maxRows: 10,
          })
          .then(
            () => {
              throw new Error('Promise unexpectedly fulfilled');
            },
            err => {
              expect(err.message).toBe('Max row count (10) exceeded');
            }
          );
      });

      it('should parse fine if the limit is not exceeded', () => {
        const workbook = new ExcelTS.Workbook();
        return workbook.xlsx.read(
          fs.createReadStream('./tests/integration/data/fibonacci.xlsx'),
          {maxRows: 20}
        );
      });
    });
  });

  describe('edit styles in existing file', () => {
    let wb: InstanceType<typeof ExcelTS.Workbook>;

    beforeEach(async () => {
      wb = new ExcelTS.Workbook();
      await wb.xlsx.readFile(
        './tests/integration/data/test-row-styles.xlsx'
      );
    });

    it('edit styles of single row instead of all', () => {
      const ws = wb.getWorksheet(1);

      ws.eachRow((row, rowNo) => {
        if (rowNo % 5 === 0) {
          row.font = {color: {argb: '00ff00'}};
        }
      });

      expect(ws.getRow(3).font.color.argb).toBe(
        ws.getRow(6).font.color.argb
      );
      expect(ws.getRow(6).font.color.argb).toBe(
        ws.getRow(9).font.color.argb
      );
      expect(ws.getRow(9).font.color.argb).toBe(
        ws.getRow(12).font.color.argb
      );
      expect(ws.getRow(12).font.color.argb).not.toBe(
        ws.getRow(15).font.color.argb
      );
      expect(ws.getRow(15).font.color.argb).not.toBe(
        ws.getRow(18).font.color.argb
      );
      expect(ws.getRow(15).font.color.argb).toBe(
        ws.getRow(10).font.color.argb
      );
      expect(ws.getRow(10).font.color.argb).toBe(
        ws.getRow(5).font.color.argb
      );
    });
  });

  describe('with a spreadsheet that contains formulas', () => {
    let worksheet: any;

    beforeAll(async () => {
      const workbook = new ExcelTS.Workbook();
      await workbook.xlsx
        .read(fs.createReadStream('./tests/integration/data/formulas.xlsx'));
      worksheet = workbook.getWorksheet();
    });

    describe('with a cell that contains a regular formula', () => {
      let cell: any;

      beforeEach(() => {
        cell = worksheet.getCell('A2');
      });

      it('should be classified as a formula cell', () => {
        expect(cell.type).toBe(ExcelTS.ValueType.Formula);
      });

      it('should have text corresponding to the evaluated formula result', () => {
        expect(cell.text).toBe('someone@example.com');
      });

      it('should have the formula source', () => {
        expect(cell.model.formula).toBe(
          '_xlfn.CONCAT("someone","@example.com")'
        );
      });
    });

    describe('with a cell that contains a hyperlinked formula', () => {
      let cell: any;

      beforeEach(() => {
        cell = worksheet.getCell('A1');
      });

      it('should be classified as a formula cell', () => {
        expect(cell.type).toBe(ExcelTS.ValueType.Hyperlink);
      });

      it('should have text corresponding to the evaluated formula result', () => {
        expect(cell.value.text).toBe('someone@example.com');
      });

      it('should have the formula source', () => {
        expect(cell.model.formula).toBe(
          '_xlfn.CONCAT("someone","@example.com")'
        );
      });

      it('should contain the linked url', () => {
        expect(cell.value.hyperlink).toBe(
          'mailto:someone@example.com'
        );
        expect(cell.hyperlink).toBe('mailto:someone@example.com');
      });
    });
  });

  describe('with a spreadsheet that contains a shared string with an escaped underscore', () => {
    let worksheet: any;

    beforeAll(async () => {
      const workbook = new ExcelTS.Workbook();
      await workbook.xlsx
        .read(
          fs.createReadStream(
            './tests/integration/data/shared_string_with_escape.xlsx'
          )
        );
      worksheet = workbook.getWorksheet();
    });

    it('should decode the underscore', () => {
      const cell = worksheet.getCell('A1');
      expect(cell.value).toBe('_x000D_');
    });
  });

  describe('with a spreadsheet that has an XML parse error in a worksheet', () => {
    let unhandledRejection;
    function unhandledRejectionHandler(err) {
      unhandledRejection = err;
    }
    beforeEach(() => {
      process.on('unhandledRejection', unhandledRejectionHandler);
    });
    afterEach(() => {
      process.removeListener('unhandledRejection', unhandledRejectionHandler);
    });

    it('should reject the promise with the sax error', () => {
      const workbook = new ExcelTS.Workbook();
      return workbook.xlsx
        .readFile('./tests/integration/data/invalid-xml.xlsx')
        .then(
          () => {
            throw new Error('Promise unexpectedly fulfilled');
          },
          err => {
            expect(err.message).toBe(
              '3:1: text data outside of root node.'
            );
            // Wait a tick before checking for an unhandled rejection
            return new Promise(setImmediate);
          }
        )
        .then(() => {
          expect(unhandledRejection).toBeUndefined();
        });
    });
  });

  describe('with a spreadsheet that is missing some files in the zip container', () => {
    it('should not break', () => {
      const workbook = new ExcelTS.Workbook();
      return workbook.xlsx.readFile(
        './tests/integration/data/missing-bits.xlsx'
      );
    });
  });

  describe('with a spreadsheet that contains images', () => {
    let worksheet: any;

    beforeAll(async () => {
      const workbook = new ExcelTS.Workbook();
      await workbook.xlsx
        .read(fs.createReadStream('./tests/integration/data/images.xlsx'));
      worksheet = workbook.getWorksheet();
    });

    describe('with image`s tl anchor', () => {
      it('Should integer part of col equals nativeCol', () => {
        worksheet.getImages().forEach(image => {
          expect(Math.floor(image.range.tl.col)).toBe(
            image.range.tl.nativeCol
          );
        });
      });
      it('Should integer part of row equals nativeRow', () => {
        worksheet.getImages().forEach(image => {
          expect(Math.floor(image.range.tl.row)).toBe(
            image.range.tl.nativeRow
          );
        });
      });
      it('Should anchor width equals to column width when custom', () => {
        const ws = worksheet;

        ws.getImages().forEach(image => {
          const col = ws.getColumn(image.range.tl.nativeCol + 1);

          if (col.isCustomWidth) {
            expect(image.range.tl.colWidth).toBe(
              Math.floor(col.width * 10000)
            );
          } else {
            expect(image.range.tl.colWidth).toBe(640000);
          }
        });
      });
      it('Should anchor height equals to row height', () => {
        const ws = worksheet;

        ws.getImages().forEach(image => {
          const row = ws.getRow(image.range.tl.nativeRow + 1);

          if (row.height) {
            expect(image.range.tl.rowHeight).toBe(
              Math.floor(row.height * 10000)
            );
          } else {
            expect(image.range.tl.rowHeight).toBe(180000);
          }
        });
      });
    });

    describe('with image`s br anchor', () => {
      it('Should integer part of col equals nativeCol', () => {
        worksheet.getImages().forEach(image => {
          expect(Math.floor(image.range.br.col)).toBe(
            image.range.br.nativeCol
          );
        });
      });
      it('Should integer part of row equals nativeRow', () => {
        worksheet.getImages().forEach(image => {
          expect(Math.floor(image.range.br.row)).toBe(
            image.range.br.nativeRow
          );
        });
      });
      it('Should anchor width equals to column width when custom', () => {
        const ws = worksheet;

        ws.getImages().forEach(image => {
          const col = ws.getColumn(image.range.br.nativeCol + 1);

          if (col.isCustomWidth) {
            expect(image.range.br.colWidth).toBe(
              Math.floor(col.width * 10000)
            );
          } else {
            expect(image.range.br.colWidth).toBe(640000);
          }
        });
      });
      it('Should anchor height equals to row height', () => {
        const ws = worksheet;

        ws.getImages().forEach(image => {
          const row = ws.getRow(image.range.br.nativeRow + 1);

          if (row.height) {
            expect(image.range.br.rowHeight).toBe(
              Math.floor(row.height * 10000)
            );
          } else {
            expect(image.range.br.rowHeight).toBe(180000);
          }
        });
      });
    });
  });
  describe('with a spreadsheet containing a defined name that kinda looks like it contains a range', () => {
    it('should not crash', () => {
      const workbook = new ExcelTS.Workbook();
      return workbook.xlsx.read(
        fs.createReadStream('./tests/integration/data/bogus-defined-name.xlsx')
      );
    });
  });
});
