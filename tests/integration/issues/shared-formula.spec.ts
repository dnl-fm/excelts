import Enums from '../../../src/doc/enums.ts';
import ExcelTS from '../../../src/index.ts';

describe('github issues', () => {
  describe('Shared Formulas', () => {
    describe('issue xyz - cells copied as a block treat formulas as values', () => {
      const explain =
        'this fails, although the cells look the same in excel. Both cells are created by copying A3:B3 to A4:F19. The first row in the new block work as espected, the rest only has values (when seen through exceljs)';
      it('copied cells should have the right formulas', () => {
        const wb = new ExcelTS.Workbook();
        return wb.xlsx
          .readFile('./tests/integration/data/fibonacci.xlsx')
          .then(() => {
            const ws = wb.getWorksheet('fib');
            expect(ws.getCell('A4').value).toEqual({
              formula: 'A3+1',
              shareType: 'shared',
              ref: 'A4:A19',
              result: 4,
            });
            expect(ws.getCell('A5').value).toEqual(
              {sharedFormula: 'A4', result: 5},
              explain
            );
          });
      });
      it('copied cells should have the right types', () => {
        const wb = new ExcelTS.Workbook();
        return wb.xlsx
          .readFile('./tests/integration/data/fibonacci.xlsx')
          .then(() => {
            const ws = wb.getWorksheet('fib');
            expect(ws.getCell('A4').type).toBe(Enums.ValueType.Formula);
            expect(ws.getCell('A5').type).toBe(Enums.ValueType.Formula);
          });
      });
      it('copied cells should have the same fields', () => {
        // to see if there are other fields on the object worth comparing
        const wb = new ExcelTS.Workbook();
        return wb.xlsx
          .readFile('./tests/integration/data/fibonacci.xlsx')
          .then(() => {
            const ws = wb.getWorksheet('fib');
            const A4 = ws.getCell('A4');
            const A5 = ws.getCell('A5');
            expect(Object.keys(A4).join()).toBe(Object.keys(A5).join());
          });
      });
    });
  });
});
