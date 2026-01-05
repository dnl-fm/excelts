
import fs from 'fs';

const exists = path => new Promise(resolve => fs.exists(path, resolve));

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
    it('should include excelts.bare.js', async () => {
      expect(await exists('./dist/excelts.bare.js')).toBe(true)
    });
    it('should include excelts.bare.min.js', async () => {
      expect(await exists('./dist/excelts.bare.min.js')).toBe(true)
    });
  });
});
