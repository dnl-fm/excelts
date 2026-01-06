import { exists } from 'fs/promises';

describe('ExcelTS', () => {
  describe('dist folder', () => {
    it('should include LICENSE', async () => {
      expect(await exists('./dist/LICENSE')).toBe(true)
    });
    it('should include excelts.js', async () => {
      expect(await exists('./dist/excelts.js')).toBe(true)
    });
    it('should include excelts.min.js', async () => {
      expect(await exists('./dist/excelts.min.js')).toBe(true)
    });
  });
});
