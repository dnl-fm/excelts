import ExcelTS from '../../../src/index.ts';

describe('github issues', () => {
  it('issue 988 - table without autofilter model', () => {
    const wb = new ExcelTS.Workbook();
    return wb.xlsx.readFile('./tests/integration/data/test-issue-988.xlsx');
  });
});
