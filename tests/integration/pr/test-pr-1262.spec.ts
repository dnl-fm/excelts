import ExcelTS from '../../../src/index.ts';

describe('github issues', () => {
  it('pull request 1262 - protect should work with streaming workbook writer', async () => {
    const workbook = new ExcelTS.stream.xlsx.WorkbookWriter({
      filename: './test.xlsx',
    });

    const sheet = workbook.addWorksheet('data');
    const row = sheet.addRow(['readonly cell']);
    row.getCell(1).protection = {
      locked: true,
    };

    expect(sheet.protect).toBeDefined();

    sheet.protect('password', {
      spinCount: 1,
    });

    await workbook.commit();

    // read in file and ensure sheetProtection is there:
    const checkBook = new ExcelTS.Workbook();
    await checkBook.xlsx.readFile('./test.xlsx');

    const checkSheet = checkBook.getWorksheet('data');
    expect(checkSheet.sheetProtection.spinCount).toBe(1);
  });
});
