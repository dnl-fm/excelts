import ExcelTS from '../../../src/index.ts';

describe('github issues', () => {
  it('issue 257 - worksheet order is not respected', () => {
    const wb = new ExcelTS.Workbook();
    return wb.xlsx
      .readFile('./tests/integration/data/test-issue-257.xlsx')
      .then(() => {
        expect(wb.worksheets.map(ws => ws.name)).toEqual([
          'First',
          'Second',
        ]);
      });
  });
});
