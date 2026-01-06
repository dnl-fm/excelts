import ExcelTS from '../../../src/index.ts';

const TEST_CSV_FILE_NAME = './tests/out/issue-995-encoding.test.csv';
const HEBREW_TEST_STRING = 'משהו שכתוב בעברית';

describe('github issues', () => {
  it('issue 995 - encoding option works fine', () => {
    const wb = new ExcelTS.Workbook();
    const ws = wb.addWorksheet('wheee');
    ws.getCell('A1').value = HEBREW_TEST_STRING;

    const writeOptions = {
      encoding: 'utf8' as BufferEncoding,
    };
    return wb.csv
      .writeFile(TEST_CSV_FILE_NAME, writeOptions)
      .then(() => {
        const ws2 = new ExcelTS.Workbook();
        return ws2.csv.readFile(TEST_CSV_FILE_NAME, {});
      })
      .then(ws2 => {
        expect(ws2.getCell('A1').value).toBe(HEBREW_TEST_STRING);
      });
  });
});
