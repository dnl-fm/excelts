// *Note*: `fs.promises` not supported before Node.js 11.14.0;
// ExcelJS version range '>=8.3.0' (as of 2023-10-08).

import { unzipSync } from 'fflate';

import ExcelTS from '../../../src/index.ts';
const PIVOT_TABLE_FILEPATHS = [
  'xl/pivotCache/pivotCacheRecords1.xml',
  'xl/pivotCache/pivotCacheDefinition1.xml',
  'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
  'xl/pivotTables/pivotTable1.xml',
  'xl/pivotTables/_rels/pivotTable1.xml.rels',
];

const TEST_XLSX_FILEPATH = './tests/out/wb.test.xlsx';

const TEST_DATA = [
  ['A', 'B', 'C', 'D', 'E'],
  ['a1', 'b1', 'c1', 4, 5],
  ['a1', 'b2', 'c1', 4, 5],
  ['a2', 'b1', 'c2', 14, 24],
  ['a2', 'b2', 'c2', 24, 35],
  ['a3', 'b1', 'c3', 34, 45],
  ['a3', 'b2', 'c3', 44, 45],
];

// =============================================================================
// Tests

describe('Workbook', () => {
  describe('Pivot Tables', () => {
    it('if pivot table added, then certain xml and rels files are added', async () => {
      const workbook = new ExcelTS.Workbook();

      const worksheet1 = workbook.addWorksheet('Sheet1');
      worksheet1.addRows(TEST_DATA);

      const worksheet2 = workbook.addWorksheet('Sheet2');
      worksheet2.addPivotTable({
        sourceSheet: worksheet1,
        rows: ['A', 'B'],
        columns: ['C'],
        values: ['E'],
        metric: 'sum',
      });

      return workbook.xlsx.writeFile(TEST_XLSX_FILEPATH).then(async () => {
        const buffer = await Bun.file(TEST_XLSX_FILEPATH).arrayBuffer();
        const files = unzipSync(new Uint8Array(buffer));
        for (const filepath of PIVOT_TABLE_FILEPATHS) {
          expect(files[filepath]).toBeDefined();
        }
      });
    });

    it('reads pivot table xml into raw containers', async () => {
      const workbook = new ExcelTS.Workbook();

      const worksheet1 = workbook.addWorksheet('Sheet1');
      worksheet1.addRows(TEST_DATA);

      const worksheet2 = workbook.addWorksheet('Sheet2');
      worksheet2.addPivotTable({
        sourceSheet: worksheet1,
        rows: ['A', 'B'],
        columns: ['C'],
        values: ['E'],
        metric: 'sum',
      });

      const buffer = await workbook.xlsx.writeBuffer();

      const readWorkbook = new ExcelTS.Workbook();
      await readWorkbook.xlsx.load(buffer);

      expect(Object.keys(readWorkbook.pivotTablesRaw)).toHaveLength(1);
      expect(Object.keys(readWorkbook.pivotCacheDefinitionsRaw)).toHaveLength(1);
      expect(Object.keys(readWorkbook.pivotCacheRecordsRaw)).toHaveLength(1);

      expect(readWorkbook.pivotTablesRaw.pivotTable1).toContain('<pivotTableDefinition');
      expect(readWorkbook.pivotCacheDefinitionsRaw.pivotCacheDefinition1).toContain('<pivotCacheDefinition');
      expect(readWorkbook.pivotCacheRecordsRaw.pivotCacheRecords1).toContain('<pivotCacheRecords');
    });

    it('if pivot table NOT added, then certain xml and rels files are not added', () => {
      const workbook = new ExcelTS.Workbook();

      const worksheet1 = workbook.addWorksheet('Sheet1');
      worksheet1.addRows(TEST_DATA);

      workbook.addWorksheet('Sheet2');

      return workbook.xlsx.writeFile(TEST_XLSX_FILEPATH).then(async () => {
        const buffer = await Bun.file(TEST_XLSX_FILEPATH).arrayBuffer();
        const files = unzipSync(new Uint8Array(buffer));
        for (const filepath of PIVOT_TABLE_FILEPATHS) {
          expect(files[filepath]).toBeUndefined();
        }
      });
    });
  });
});
