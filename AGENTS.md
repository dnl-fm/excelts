# ExcelTS - Agent Guide

Bun-first TypeScript fork of ExcelJS for reading/writing Excel XLSX and CSV files.

## Quick Reference

```bash
bun install              # Install dependencies
bun run build            # Build browser bundles
bun test tests/unit      # Run unit tests
bun test tests/integration  # Run integration tests
bun run lint             # Run oxlint
bun run format:write     # Auto-format with oxfmt
```

## Structure

```
excelts/
├── src/
│   ├── doc/                 # Core document models
│   │   ├── workbook.ts      # Main Workbook class
│   │   ├── worksheet.ts     # Worksheet with cells, rows, columns
│   │   ├── cell.ts          # Cell values, formulas, styles
│   │   ├── row.ts           # Row operations
│   │   ├── column.ts        # Column definitions
│   │   └── enums.ts         # ValueType, FormulaType, etc.
│   ├── xlsx/
│   │   ├── xlsx.ts          # XLSX read/write orchestration
│   │   └── xform/           # XML transformers (serialize/parse)
│   │       ├── base-xform.ts    # Base class for all xforms
│   │       ├── book/            # Workbook XML parts
│   │       ├── sheet/           # Worksheet XML parts
│   │       ├── style/           # Styles (fonts, fills, borders)
│   │       ├── strings/         # Shared strings table
│   │       ├── drawing/         # Images, charts
│   │       ├── table/           # Table definitions
│   │       └── comment/         # Cell comments
│   ├── csv/                 # CSV read/write
│   ├── stream/              # Streaming read/write
│   │   └── xlsx/
│   │       ├── workbook-writer.ts
│   │       └── workbook-reader.ts
│   ├── utils/               # Utilities
│   │   ├── col-cache.ts     # Column letter ↔ number conversion
│   │   ├── xml-stream.ts    # XML writer
│   │   ├── parse-sax.ts     # SAX XML parser
│   │   └── utils.ts         # General helpers
│   └── exceljs.nodejs.ts    # Main entry point
├── tests/
│   ├── unit/                # Unit tests per module
│   ├── integration/         # Full read/write tests
│   ├── utils/               # Test helpers, fixtures
│   └── preload.ts           # Sets up `verquire` global
├── dist/                    # Built browser bundles (gitignored)
└── index.d.ts               # TypeScript declarations
```

## Tech Stack

- **Runtime**: Bun 1.3+ (required)
- **Language**: TypeScript (ESM, `.ts` imports)
- **Build**: `bun build` for browser bundles
- **Test**: Bun's native test runner
- **Lint**: oxlint
- **Format**: oxfmt

## Key Concepts

### Xform Pattern
XML transformers handle bidirectional conversion between JavaScript models and Excel XML:

```typescript
class MyXform extends BaseXform {
  // Model → XML
  render(xmlStream, model) {
    xmlStream.openNode('tag');
    xmlStream.closeNode();
  }
  
  // XML → Model (SAX-based)
  parseOpen(node) { /* element start */ }
  parseText(text) { /* text content */ }
  parseClose(name) { /* element end, return false when done */ }
}
```

### Document Hierarchy
```
Workbook
  └── Worksheet[]
        ├── Row[]
        │     └── Cell[]
        ├── Column[]
        ├── MergeCells
        ├── DataValidations
        └── ConditionalFormatting
```

### Value Types
```typescript
// src/doc/enums.ts
ValueType.Null        // 0 - empty cell
ValueType.Number      // 2 - numeric
ValueType.String      // 3 - text
ValueType.Date        // 4 - date/time
ValueType.Formula     // 6 - formula
ValueType.RichText    // 8 - formatted text
ValueType.Boolean     // 9 - true/false
ValueType.Error       // 10 - #N/A, #REF!, etc.
```

## Test Patterns

### Using verquire
Tests use `verquire()` to load source modules (set up in `tests/preload.ts`):

```typescript
const Excel = verquire('exceljs');
const Workbook = verquire('doc/workbook');
const CellXform = verquire('xlsx/xform/sheet/cell-xform');
```

### Test Structure
```typescript
import { describe, it, expect } from 'bun:test';

describe('MyFeature', () => {
  it('should do something', () => {
    const wb = new Excel.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 42;
    expect(ws.getCell('A1').value).toBe(42);
  });
});
```

### Xform Tests
Unit tests for xforms follow a pattern with expectations:

```typescript
const expectations = [
  {
    title: 'Basic Case',
    create: () => new MyXform(),
    preparedModel: { /* input model */ },
    xml: '<expected>xml</expected>',
    parsedModel: { /* expected parse result */ },
    tests: ['render', 'parse'],
  },
];

testXformHelper(expectations);
```

## Common Tasks

### Add New Cell Feature
1. Update `src/doc/cell.ts` with property/method
2. Update `src/xlsx/xform/sheet/cell-xform.ts` for XML handling
3. Add tests in `tests/unit/xlsx/xform/sheet/cell-xform.spec.ts`

### Add New Worksheet Feature
1. Update `src/doc/worksheet.ts`
2. Update `src/xlsx/xform/sheet/worksheet-xform.ts`
3. May need new child xform in `src/xlsx/xform/sheet/`

### Add New Style
1. Check `src/xlsx/xform/style/` for relevant xform
2. Update styles-xform.ts if adding new style category
3. Test with `tests/unit/xlsx/xform/style/`

## File Naming

- Source: `kebab-case.ts` (e.g., `cell-xform.ts`)
- Tests: `kebab-case.spec.ts` (e.g., `cell-xform.spec.ts`)
- Test data: `kebab-case.json` or `name.N.M.json` for versioned fixtures

## Gotchas

1. **XML Attribute Order**: Excel is sensitive to attribute order in some elements. Tests may fail on ordering differences even if XML is semantically equivalent.

2. **Date Handling**: Excel stores dates as numbers (days since 1900). Use `utils.dateToExcel()` / `utils.excelToDate()`.

3. **Shared Strings**: Strings are stored in a shared table for deduplication. The `SharedStringsXform` manages this.

4. **Style IDs**: Styles are referenced by numeric ID. The `StylesXform` maintains the mapping.

5. **Streaming vs DOM**: 
   - `Workbook` loads entire file into memory
   - `WorkbookWriter`/`WorkbookReader` stream for large files

6. **verquire paths**: Use forward slashes, no `.ts` extension, relative to `src/`:
   ```typescript
   verquire('xlsx/xform/sheet/cell-xform')  // ✓
   verquire('./src/xlsx/xform/sheet/cell-xform.ts')  // ✗
   ```

## Dependencies

**Production**: fflate, dayjs, saxes, udsv

**Dev**: @types/bun, express, oxfmt, oxlint, typescript

**Removed** (replaced with Bun-native solutions):
- `uuid` → `crypto.randomUUID()` (Web Crypto)
- `readable-stream` → Custom async iterables + SimpleBuffer
- `tmp` → `Bun.env.TMPDIR` + manual cleanup
- `fast-csv` → `udsv` (10x faster, no Node streams)
- `Buffer` → `Uint8Array` + `src/utils/bytes.ts` helpers
- `jszip/archiver/unzipper` → `fflate` (smaller, faster, browser-native)

**Node compatibility**: None required. Fully Bun-native.

## Status

- ✅ ESM migration complete
- ✅ Bun-native test runner
- ✅ TypeScript throughout
- ✅ Dependency cleanup (uuid, readable-stream, tmp, fast-csv removed)
- ✅ Main xlsx path is Bun-native (no Node stream/fs dependencies)
- ✅ CSV module is Bun-native (uses udsv, Bun.file)
- ✅ Streaming module uses Bun file I/O with Node stream bridge for ZIP libs
- ✅ Created SimpleBuffer, SimpleEventEmitter for Bun-native code
- ✅ Replaced Buffer with Uint8Array throughout (`src/utils/bytes.ts`)
- ✅ Refactored `stream-buf.ts` to be browser-native (uses SimpleEventEmitter + bytes.ts)
- ✅ Replaced jszip/archiver/unzipper with fflate
- ✅ WorkbookWriter now uses Bun.write() for file output (removed Node stream option)
- ✅ All 1079 tests passing (883 unit + 196 integration)

## TODO

- [ ] Generate `index.d.ts` from source
