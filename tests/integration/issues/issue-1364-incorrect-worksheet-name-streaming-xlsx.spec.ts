import ExcelTS from '../../../src/index.ts';

const TEST_XLSX_FILE_NAME = './tests/integration/data/test-issue-1364.xlsx';

describe('github issues', () => {
  it('issue 1364 - Incorrect Worksheet Name on Streaming XLSX Reader', async () => {
    const workbookReader = new ExcelTS.stream.xlsx.WorkbookReader(
      TEST_XLSX_FILE_NAME,
      {}
    );
    workbookReader.read();
    workbookReader.on('worksheet', (worksheet: {name: string}) => {
      expect(worksheet.name).toBe('Sum Worksheet');
    });
  });
});
