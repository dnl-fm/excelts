import fs from 'fs';

import ExcelTS from '../../../src/index.ts';
describe('github issues: Date field with cache style', () => {
  const rows = [];
  beforeEach(
    () =>
      new Promise((resolve, reject) => {
        const workbookReader = new ExcelTS.stream.xlsx.WorkbookReader(
          fs.createReadStream('./tests/integration/data/dateIssue.xlsx'),
          {
            worksheets: 'emit',
            styles: 'cache',
            sharedStrings: 'cache',
            hyperlinks: 'ignore',
            entries: 'ignore',
          }
        );
        workbookReader.read();
        workbookReader.on('worksheet', worksheet =>
          worksheet.on('row', row => rows.push(row.values[1]))
        );
        workbookReader.on('end', resolve);
        workbookReader.on('error', reject);
      })
  );
  it('issue 1328 - should emit row with Date Object', () => {
    expect(rows).toEqual([
      'Date',
      new Date('2020-11-20T00:00:00.000Z'),
    ]);
  });
});
