import tools from '../../../utils/tools.ts';
import sheetPropertiesJson from '../../../utils/data/sheet-properties.json';
import pageSetupJson from '../../../utils/data/page-setup.json';

import Excel from '../../../../src/index.ts';
process.env.EXCEL_NATIVE = 'yes';

const TEST_XLSX_FILE_NAME = './tests/out/wb.test.xlsx';
const RT_ARR = [
  {text: 'First Line:\n', font: {bold: true}},
  {text: 'Second Line\n'},
  {text: 'Third Line\n'},
  {text: 'Last Line'},
];
const TEST_VALUE = {
  richText: RT_ARR,
};
const TEST_NOTE = {
  texts: RT_ARR,
};

describe('pr related issues', () => {
  describe('pr 896 add xml:space="preserve" for all whitespaces', () => {
    it('should store cell text and comment with leading new line', () => {
      const properties = tools.fix(sheetPropertiesJson);
      const pageSetup = tools.fix(pageSetupJson);

      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('sheet1', {
        properties,
        pageSetup,
      });

      ws.getColumn(1).width = 20;
      ws.getCell('A1').value = TEST_VALUE;
      ws.getCell('A1').note = TEST_NOTE;
      ws.getCell('A1').alignment = {wrapText: true};

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(wb2 => {
          const ws2 = wb2.getWorksheet('sheet1');
          expect(ws2).toBeDefined();
          expect(ws2.getCell('A1').value).toEqual(TEST_VALUE);
        });
    });
  });
});
