import { describe, it, expect } from 'bun:test';
import ExcelTS from '../../../src/index.ts';

describe('Text Format Detection', () => {
  describe('Cell Font Properties', () => {
    it('detects bold text', () => {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('Sheet1');

      ws.getCell('A1').value = 'Bold Text';
      ws.getCell('A1').font = { bold: true };

      expect(ws.getCell('A1').font).toMatchObject({ bold: true });
    });

    it('detects italic text', () => {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('Sheet1');

      ws.getCell('A1').value = 'Italic Text';
      ws.getCell('A1').font = { italic: true };

      expect(ws.getCell('A1').font).toMatchObject({ italic: true });
    });

    it('detects underline styles', () => {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('Sheet1');

      // Simple underline (boolean)
      ws.getCell('A1').value = 'Underlined';
      ws.getCell('A1').font = { underline: true };
      expect(ws.getCell('A1').font).toMatchObject({ underline: true });

      // Single underline
      ws.getCell('A2').value = 'Single Underline';
      ws.getCell('A2').font = { underline: 'single' };
      expect(ws.getCell('A2').font).toMatchObject({ underline: 'single' });

      // Double underline
      ws.getCell('A3').value = 'Double Underline';
      ws.getCell('A3').font = { underline: 'double' };
      expect(ws.getCell('A3').font).toMatchObject({ underline: 'double' });
    });

    it('detects strikethrough text', () => {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('Sheet1');

      ws.getCell('A1').value = 'Strikethrough';
      ws.getCell('A1').font = { strike: true };

      expect(ws.getCell('A1').font).toMatchObject({ strike: true });
    });

    it('detects font color (ARGB)', () => {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('Sheet1');

      ws.getCell('A1').value = 'Red Text';
      ws.getCell('A1').font = { color: { argb: 'FFFF0000' } };

      expect(ws.getCell('A1').font).toMatchObject({
        color: { argb: 'FFFF0000' },
      });
    });

    it('detects font color (theme)', () => {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('Sheet1');

      ws.getCell('A1').value = 'Theme Color';
      ws.getCell('A1').font = { color: { theme: 4 } };

      expect(ws.getCell('A1').font).toMatchObject({
        color: { theme: 4 },
      });
    });

    it('detects font name', () => {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('Sheet1');

      ws.getCell('A1').value = 'Arial Text';
      ws.getCell('A1').font = { name: 'Arial' };

      expect(ws.getCell('A1').font).toMatchObject({ name: 'Arial' });
    });

    it('detects font size', () => {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('Sheet1');

      ws.getCell('A1').value = 'Large Text';
      ws.getCell('A1').font = { size: 24 };

      expect(ws.getCell('A1').font).toMatchObject({ size: 24 });
    });

    it('detects combined font properties', () => {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('Sheet1');

      const complexFont = {
        name: 'Arial',
        size: 14,
        bold: true,
        italic: true,
        underline: 'double' as const,
        strike: true,
        color: { argb: 'FF0000FF' },
      };

      ws.getCell('A1').value = 'Complex Formatting';
      ws.getCell('A1').font = complexFont;

      expect(ws.getCell('A1').font).toMatchObject(complexFont);
    });

    it('detects superscript and subscript', () => {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('Sheet1');

      ws.getCell('A1').value = 'Superscript';
      ws.getCell('A1').font = { vertAlign: 'superscript' };
      expect(ws.getCell('A1').font).toMatchObject({ vertAlign: 'superscript' });

      ws.getCell('A2').value = 'Subscript';
      ws.getCell('A2').font = { vertAlign: 'subscript' };
      expect(ws.getCell('A2').font).toMatchObject({ vertAlign: 'subscript' });
    });

    it('detects outline font', () => {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('Sheet1');

      ws.getCell('A1').value = 'Outline';
      ws.getCell('A1').font = { outline: true };

      expect(ws.getCell('A1').font).toMatchObject({ outline: true });
    });
  });

  describe('Rich Text (In-Cell Formatting)', () => {
    it('detects rich text with bold segment', () => {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('Sheet1');

      const richText = {
        richText: [
          { text: 'Normal ' },
          { text: 'Bold', font: { bold: true } },
          { text: ' Normal' },
        ],
      };

      ws.getCell('A1').value = richText;

      expect(ws.getCell('A1').type).toBe(ExcelTS.ValueType.RichText);
      expect(ws.getCell('A1').value).toEqual(richText);
    });

    it('detects rich text with multiple format segments', () => {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('Sheet1');

      const richText = {
        richText: [
          { text: 'Bold', font: { bold: true } },
          { text: ' and ' },
          { text: 'Italic', font: { italic: true } },
          { text: ' and ' },
          { text: 'Red', font: { color: { argb: 'FFFF0000' } } },
        ],
      };

      ws.getCell('A1').value = richText;

      expect(ws.getCell('A1').type).toBe(ExcelTS.ValueType.RichText);
      const value = ws.getCell('A1').value as { richText: Array<{ text: string; font?: Record<string, unknown> }> };
      expect(value.richText).toHaveLength(5);
      expect(value.richText[0].font).toMatchObject({ bold: true });
      expect(value.richText[2].font).toMatchObject({ italic: true });
      expect(value.richText[4].font).toMatchObject({ color: { argb: 'FFFF0000' } });
    });

    it('extracts plain text from rich text', () => {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('Sheet1');

      ws.getCell('A1').value = {
        richText: [
          { text: 'Hello ' },
          { text: 'World', font: { bold: true } },
          { text: '!' },
        ],
      };

      // The .text property should return concatenated plain text
      expect(ws.getCell('A1').text).toBe('Hello World!');
    });

    it('detects rich text with complex segment formatting', () => {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('Sheet1');

      const richText = {
        richText: [
          {
            text: 'Complex',
            font: {
              name: 'Arial',
              size: 12,
              bold: true,
              italic: true,
              underline: 'single' as const,
              color: { argb: 'FF0000FF' },
            },
          },
        ],
      };

      ws.getCell('A1').value = richText;

      const value = ws.getCell('A1').value as { richText: Array<{ text: string; font?: Record<string, unknown> }> };
      expect(value.richText[0].font).toMatchObject({
        name: 'Arial',
        size: 12,
        bold: true,
        italic: true,
        underline: 'single',
        color: { argb: 'FF0000FF' },
      });
    });
  });

  describe('Round-Trip Format Preservation', () => {
    it('preserves cell font through write/read cycle', async () => {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('Sheet1');

      const testFont = {
        name: 'Arial',
        size: 14,
        bold: true,
        italic: true,
        color: { argb: 'FFFF0000' },
      };

      ws.getCell('A1').value = 'Formatted Text';
      ws.getCell('A1').font = testFont;

      // Write to buffer and read back
      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new ExcelTS.Workbook();
      await wb2.xlsx.load(buffer as Buffer);

      const ws2 = wb2.getWorksheet('Sheet1')!;
      expect(ws2.getCell('A1').value).toBe('Formatted Text');
      expect(ws2.getCell('A1').font).toMatchObject(testFont);
    });

    it('preserves rich text through write/read cycle', async () => {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('Sheet1');

      const richText = {
        richText: [
          { text: 'Normal ' },
          { text: 'Bold', font: { bold: true } },
          { text: ' ' },
          { text: 'Italic', font: { italic: true } },
        ],
      };

      ws.getCell('A1').value = richText;

      // Write to buffer and read back
      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new ExcelTS.Workbook();
      await wb2.xlsx.load(buffer as Buffer);

      const ws2 = wb2.getWorksheet('Sheet1')!;
      expect(ws2.getCell('A1').type).toBe(ExcelTS.ValueType.RichText);
      expect(ws2.getCell('A1').text).toBe('Normal Bold Italic');

      const value = ws2.getCell('A1').value as { richText: Array<{ text: string; font?: Record<string, unknown> }> };
      expect(value.richText).toHaveLength(4);
      expect(value.richText[1].font).toMatchObject({ bold: true });
      expect(value.richText[3].font).toMatchObject({ italic: true });
    });

    it('preserves multiple formatted cells through write/read cycle', async () => {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('Sheet1');

      // Set up various formatted cells
      ws.getCell('A1').value = 'Bold';
      ws.getCell('A1').font = { bold: true };

      ws.getCell('A2').value = 'Italic';
      ws.getCell('A2').font = { italic: true };

      ws.getCell('A3').value = 'Red';
      ws.getCell('A3').font = { color: { argb: 'FFFF0000' } };

      ws.getCell('A4').value = 'Large';
      ws.getCell('A4').font = { size: 24 };

      ws.getCell('A5').value = 'Underline';
      ws.getCell('A5').font = { underline: true };

      // Write and read back
      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new ExcelTS.Workbook();
      await wb2.xlsx.load(buffer as Buffer);

      const ws2 = wb2.getWorksheet('Sheet1')!;

      expect(ws2.getCell('A1').font).toMatchObject({ bold: true });
      expect(ws2.getCell('A2').font).toMatchObject({ italic: true });
      expect(ws2.getCell('A3').font).toMatchObject({ color: { argb: 'FFFF0000' } });
      expect(ws2.getCell('A4').font).toMatchObject({ size: 24 });
      expect(ws2.getCell('A5').font).toMatchObject({ underline: true });
    });
  });

  describe('Format Detection Helpers', () => {
    it('can check if cell has any formatting', () => {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('Sheet1');

      ws.getCell('A1').value = 'Plain';
      ws.getCell('A2').value = 'Formatted';
      ws.getCell('A2').font = { bold: true };

      // Helper function to check if cell has font formatting
      const hasFormatting = (cell: ExcelTS.Cell) => {
        const font = cell.font as Record<string, unknown> | undefined;
        if (!font) return false;
        return Object.keys(font).length > 0;
      };

      expect(hasFormatting(ws.getCell('A1'))).toBe(false);
      expect(hasFormatting(ws.getCell('A2'))).toBe(true);
    });

    it('can check specific format properties', () => {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('Sheet1');

      ws.getCell('A1').value = 'Bold';
      ws.getCell('A1').font = { bold: true };

      ws.getCell('A2').value = 'Not Bold';
      ws.getCell('A2').font = { italic: true };

      // Helper to check if bold
      const isBold = (cell: ExcelTS.Cell) => {
        const font = cell.font as { bold?: boolean } | undefined;
        return font?.bold === true;
      };

      expect(isBold(ws.getCell('A1'))).toBe(true);
      expect(isBold(ws.getCell('A2'))).toBe(false);
    });

    it('can extract all formatting from a cell', () => {
      const wb = new ExcelTS.Workbook();
      const ws = wb.addWorksheet('Sheet1');

      const testFont = {
        name: 'Verdana',
        size: 16,
        bold: true,
        italic: true,
        underline: 'double' as const,
        strike: true,
        color: { argb: 'FF00FF00' },
      };

      ws.getCell('A1').value = 'All Formats';
      ws.getCell('A1').font = testFont;

      // Extract and verify each property
      const font = ws.getCell('A1').font as typeof testFont;
      expect(font.name).toBe('Verdana');
      expect(font.size).toBe(16);
      expect(font.bold).toBe(true);
      expect(font.italic).toBe(true);
      expect(font.underline).toBe('double');
      expect(font.strike).toBe(true);
      expect(font.color).toEqual({ argb: 'FF00FF00' });
    });
  });
});
