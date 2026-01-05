import {expect, describe, it, beforeAll} from 'bun:test';

const ExcelJS = verquire('exceljs');

/**
 * Compare dates ignoring milliseconds
 */
function expectDateEqual(actual: Date, expected: Date): void {
  expect(actual.getFullYear()).toBe(expected.getFullYear());
  expect(actual.getMonth()).toBe(expected.getMonth());
  expect(actual.getDate()).toBe(expected.getDate());
}

// =============================================================================
// This spec is based around a gold standard Excel workbook 'gold.xlsx'

describe('Gold Book', () => {
  describe('Read', () => {
    let wb: any;

    beforeAll(async () => {
      wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(`${import.meta.dir}/data/gold.xlsx`);
    });

    it('Values', () => {
      const ws = wb.getWorksheet('Values');

      expect(ws.getCell('B1').value).toBe('I am Text');
      expect(ws.getCell('B2').value).toBe(3.14);
      expect(ws.getCell('B3').value).toBe(5);

      expectDateEqual(
        ws.getCell('B4').value,
        new Date('2016-05-17T00:00:00.000Z')
      );

      expect(ws.getCell('B5').value).toEqual({
        formula: 'B1',
        result: 'I am Text',
      });

      expect(ws.getCell('B6').value).toEqual({
        hyperlink: 'https://www.npmjs.com/package/exceljs',
        text: 'exceljs',
      });

      expect(ws.lastColumn).toBe(ws.getColumn(2));
      expect(ws.lastRow).toBe(ws.getRow(6));
    });

    it('Styles', () => {});
  });
});
