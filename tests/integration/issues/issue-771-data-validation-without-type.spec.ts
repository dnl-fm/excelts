import ExcelTS from '../../../src/index.ts';

describe('github issues', () => {
  it('issue 771 - Issue with dataValidation without type and with formula1 or formula2', () => {
    const wb = new ExcelTS.Workbook();
    return wb.xlsx.readFile('./tests/integration/data/test-issue-771.xlsx');
  });
});
