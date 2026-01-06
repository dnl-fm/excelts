# ExcelTS

Read and write Excel files. TypeScript, minimal dependencies, works in Bun and browsers.

## Why?

[ExcelJS](https://github.com/exceljs/exceljs) is great but carries Node.js baggage. ExcelTS is a modernized fork:

| | ExcelJS | ExcelTS |
|---|---------|---------|
| Runtime | Node.js | Bun / Browser |
| Modules | CJS + ESM | ESM only |
| ZIP handling | 3 packages (170KB) | fflate (8KB) |
| Bundle size | 2.1 MB | 289 KB |
| Node polyfills | Required | None |
| Types | Separate `.d.ts` | Native TypeScript |

**Browser-native**: Uses `Uint8Array`, `TextEncoder/TextDecoder`, and Web Crypto APIs. No Buffer polyfill needed.

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

sheet.columns = [
  { header: 'Product', key: 'product', width: 20 },
  { header: 'Revenue', key: 'revenue', width: 15 },
];

sheet.addRow({ product: 'Widget', revenue: 1500 });
sheet.addRow({ product: 'Gadget', revenue: 2300 });

sheet.getRow(1).font = { bold: true };

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

### Browser: Load from ArrayBuffer

```html
<script src="excelts.min.js"></script>
<script>
async function loadExcel(arrayBuffer) {
  const workbook = new ExcelTS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  
  workbook.eachSheet((sheet) => {
    console.log(sheet.name);
    sheet.eachRow((row) => {
      console.log(row.values);
    });
  });
}

// From fetch
const response = await fetch('data.xlsx');
const buffer = await response.arrayBuffer();
await loadExcel(buffer);

// From file input
input.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  const buffer = await file.arrayBuffer();
  await loadExcel(buffer);
});

// From drag & drop
dropzone.addEventListener('drop', async (e) => {
  const file = e.dataTransfer.files[0];
  const buffer = await file.arrayBuffer();
  await loadExcel(buffer);
});
</script>
```

### Browser: Generate and download

```javascript
const workbook = new ExcelTS.Workbook();
const sheet = workbook.addWorksheet('Export');

sheet.addRow(['Name', 'Email', 'Date']);
sheet.addRow(['John', 'john@example.com', new Date()]);

// One-liner download
await workbook.xlsx.download('export.xlsx');

// Or get Blob for custom handling (uploads, etc.)
const blob = await workbook.xlsx.writeBlob();
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

## Browser Build

The browser build (`dist/excelts.min.js`) uses only native browser APIs:

- `Uint8Array` instead of Node.js Buffer
- `TextEncoder` / `TextDecoder` for string encoding
- `crypto.randomUUID()` and `crypto.getRandomValues()` for crypto
- `fflate` for ZIP compression (pure JS, no WASM)

No polyfills required. Works in all modern browsers.

## API

Full API matches [ExcelJS](https://github.com/exceljs/exceljs#interface). Key classes:

- `Workbook` - Container for worksheets
- `Worksheet` - Grid of cells, rows, columns
- `Row` / `Cell` - Data and formatting
- `WorkbookWriter` / `WorkbookReader` - Streaming API

## License

MIT

Fork of [ExcelJS](https://github.com/exceljs/exceljs) by Guyon Roche.
