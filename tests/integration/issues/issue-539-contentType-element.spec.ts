import ExcelTS from '../../../src/index.ts';

describe('github issues', () => {
  describe('issue 539 - <contentType /> element', () => {
    it('Reading 1904.xlsx', () => {
      const wb = new ExcelTS.Workbook();
      return wb.xlsx.readFile(
        './tests/integration/data/1519293514-KRISHNAPATNAM_LINE_UP.xlsx'
      );
    });
  });
});
