
import fs from 'fs';

import ExcelTS from '../../../src/index.ts';
// this file to contain integration tests created from github issues
const TEST_XLSX_FILE_NAME = './tests/out/wb-issue-880.test.xlsx';

describe('github issues', () => {
  it('issue 880 - malformed comment crashes on write', () => {
    const wb = new ExcelTS.Workbook();
    return wb.xlsx
      .readFile('./tests/integration/data/test-issue-880.xlsx')
      .then(() => {
        wb.xlsx
          .writeBuffer({
            useStyles: true,
            useSharedStrings: true,
          })
          .then(function(buffer) {
            const wstream = fs.createWriteStream(TEST_XLSX_FILE_NAME);
            wstream.write(buffer);
            wstream.end();
          });
      });
  });
});
