import ExcelTS from '../../../src/index.ts';

describe('github issues', () => {
  describe('pull request 1487 - lastColumn with an empty column', () => {
    it('Reading 1904.xlsx', () => {
      const wb = new ExcelTS.Workbook();
      return wb.xlsx.readFile('./tests/integration/data/1904.xlsx').then(() => {
        const ws = wb.getWorksheet('Sheet1');
        expect(ws.lastColumn).toBe(ws.getColumn(2));
      });
    });
  });
  // the new property is also tested in gold.spec.js
});
