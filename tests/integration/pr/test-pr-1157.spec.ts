import ExcelTS from '../../../src/index.ts';

const TEST_XLSX_FILE_NAME = './tests/out/wb.test.xlsx';

describe('github issues', () => {
  it('pull request 1204 - Read and write data validation should be successful', async () => {
    const wb = new ExcelTS.Workbook();
    await wb.xlsx.readFile('./tests/integration/data/test-pr-1204.xlsx');
    const expected = {
      E1: {
        type: 'textLength',
        formulae: [2],
        showInputMessage: true,
        showErrorMessage: true,
        operator: 'greaterThan',
      },
      E4: {
        type: 'textLength',
        formulae: [2],
        showInputMessage: true,
        showErrorMessage: true,
        operator: 'greaterThan',
      },
    };
    const ws = wb.getWorksheet(1);
    expect(ws.dataValidations.model).toEqual(expected);
    await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
  });
});
