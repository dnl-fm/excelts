import ExcelTS from '../../../src/index.ts';

describe('github issues', () => {
  it('issue 176 - Unexpected xml node in parseOpen', () => {
    const wb = new ExcelTS.Workbook();
    return wb.xlsx
      .readFile('./tests/integration/data/test-issue-176.xlsx')
      .then(() => {
        // arriving here is success
        expect(true).toBe(true);
      });
  });
});
