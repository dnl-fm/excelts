import ExcelTS from '../../../src/index.ts';

describe('github issues', () => {
  it('issue 1669 - optional autofilter and custom autofilter on tables', () => {
    const wb = new ExcelTS.Workbook();
    return wb.xlsx.readFile('./tests/integration/data/test-issue-1669.xlsx');
  });
});
