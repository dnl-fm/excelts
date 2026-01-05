import ExcelTS from '../../../src/index.ts';

describe('github issues', () => {
  it('pull request 1220 - The worksheet should not be undefined', async () => {
    const wb = new ExcelTS.Workbook();
    await wb.xlsx.readFile('./tests/integration/data/test-pr-1220.xlsx');
    const ws = wb.getWorksheet(1);
    expect(ws).not.toBe(undefined);
  });
});
