import ExcelTS from '../../../src/index.ts';

describe('github issues', () => {
  it('issue 163 - Error while using xslx readFile method', () => {
    const wb = new ExcelTS.Workbook();
    return wb.xlsx
      .readFile('./tests/integration/data/test-issue-163.xlsx')
      .then(() => {
        // arriving here is success
        expect(true).toBe(true);
      });
  });
});
