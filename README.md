# ExcelTS

[![Build Status](https://github.com/dnl-fm/excelts/actions/workflows/tests.yml/badge.svg?branch=main&event=push)](https://github.com/dnl-fm/excelts/actions/workflows/tests.yml)

> Bun-first TypeScript fork of [ExcelJS](https://github.com/exceljs/exceljs) for reading/writing Excel XLSX and CSV files.

## Features

- Read and write XLSX files
- Streaming API for large files
- Rich cell formatting (fonts, colors, borders, fills)
- Formulas, hyperlinks, images
- Data validations and conditional formatting
- CSV import/export

## Installation

```bash
bun add excelts
```

## Quick Start

```typescript
import ExcelJS from 'excelts';

// Create a workbook
const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet('Sheet1');

// Add data
sheet.columns = [
  { header: 'Name', key: 'name', width: 20 },
  { header: 'Age', key: 'age', width: 10 },
];

sheet.addRow({ name: 'John', age: 30 });
sheet.addRow({ name: 'Jane', age: 25 });

// Save to file
await workbook.xlsx.writeFile('output.xlsx');
```

## Reading Files

```typescript
const workbook = new ExcelJS.Workbook();
await workbook.xlsx.readFile('input.xlsx');

workbook.eachSheet((sheet, id) => {
  sheet.eachRow((row, rowNumber) => {
    console.log(row.values);
  });
});
```

## Streaming (Large Files)

```typescript
// Writing
const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
  filename: 'large.xlsx',
});
const sheet = workbook.addWorksheet('Data');
for (let i = 0; i < 1000000; i++) {
  sheet.addRow([i, `Row ${i}`]).commit();
}
await workbook.commit();

// Reading
const workbook = new ExcelJS.stream.xlsx.WorkbookReader('large.xlsx');
for await (const sheet of workbook) {
  for await (const row of sheet) {
    console.log(row.values);
  }
}
```

## Changes from ExcelJS

- **Bun-first**: Uses Bun for runtime, testing, and building
- **TypeScript native**: All source code in TypeScript
- **ESM only**: No CommonJS, no ES5 builds
- **Modern tooling**: oxlint, oxfmt, bun test
- **Simplified deps**: Removed legacy polyfills and build tools

## Development

```bash
bun install          # Install dependencies
bun test             # Run all tests
bun run build        # Build browser bundles
bun run lint         # Run linter
bun run format:write # Auto-format
```

## License

MIT - Originally created by [Guyon Roche](https://github.com/guyonroche)

## Credits

This is a modernized fork of [ExcelJS](https://github.com/exceljs/exceljs). All credit for the core functionality goes to the original authors and contributors.
