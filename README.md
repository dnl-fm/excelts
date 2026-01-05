# ExcelTS

Read and write Excel files. TypeScript, minimal dependencies, works in Bun and browsers.

## Why?

[ExcelJS](https://github.com/exceljs/exceljs) is great but carries Node.js baggage. ExcelTS is a modernized fork:

| | ExcelJS | ExcelTS |
|---|---------|---------|
| Runtime | Node.js | Bun / Browser |
| Modules | CJS + ESM | ESM only |
| ZIP handling | 3 packages (170KB) | fflate (8KB) |
| Types | Separate `.d.ts` | Native TypeScript |

## Install

```bash
bun add excelts
```

## Examples

### Write a spreadsheet

```typescript
import ExcelTS from 'excelts';

const workbook = new ExcelTS.Workbook();
const sheet = workbook.addWorksheet('Sales');

// Define columns
sheet.columns = [
  { header: 'Product', key: 'product', width: 20 },
  { header: 'Revenue', key: 'revenue', width: 15 },
];

// Add rows
sheet.addRow({ product: 'Widget', revenue: 1500 });
sheet.addRow({ product: 'Gadget', revenue: 2300 });

// Style the header
sheet.getRow(1).font = { bold: true };

// Save
await workbook.xlsx.writeFile('sales.xlsx');
```

### Read a spreadsheet

```typescript
const workbook = new ExcelTS.Workbook();
await workbook.xlsx.readFile('sales.xlsx');

const sheet = workbook.getWorksheet('Sales');
sheet.eachRow((row, num) => {
  console.log(num, row.values);
});
```

### Stream large files

```typescript
// Write 1M rows without memory issues
const writer = new ExcelTS.stream.xlsx.WorkbookWriter({ filename: 'big.xlsx' });
const sheet = writer.addWorksheet('Data');

for (let i = 0; i < 1_000_000; i++) {
  sheet.addRow([i, `Row ${i}`]).commit();
}
await writer.commit();

// Read large files
const reader = new ExcelTS.stream.xlsx.WorkbookReader('big.xlsx');
for await (const sheet of reader) {
  for await (const row of sheet) {
    // Process row
  }
}
```

### Cell formatting

```typescript
const cell = sheet.getCell('A1');

cell.value = 1234.56;
cell.numFmt = '$#,##0.00';
cell.font = { bold: true, color: { argb: 'FF0000' } };
cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF00' } };
cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
```

### Formulas

```typescript
sheet.getCell('A1').value = 10;
sheet.getCell('A2').value = 20;
sheet.getCell('A3').value = { formula: 'SUM(A1:A2)' };
```

### Images

```typescript
const imageId = workbook.addImage({
  filename: 'logo.png',
  extension: 'png',
});

sheet.addImage(imageId, 'B2:D6');
```

## API

Full API matches [ExcelJS](https://github.com/exceljs/exceljs#interface). Key classes:

- `Workbook` - Container for worksheets
- `Worksheet` - Grid of cells, rows, columns
- `Row` / `Cell` - Data and formatting
- `WorkbookWriter` / `WorkbookReader` - Streaming API

## License

MIT

Fork of [ExcelJS](https://github.com/exceljs/exceljs) by Guyon Roche.
