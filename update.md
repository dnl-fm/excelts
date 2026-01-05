# ExcelTS Update Tasks

## JSR Publishing: Add Explicit Return Types

JSR requires explicit return types for all public API functions. Currently ~487 errors across 14 files.

**Status:** Using `--allow-slow-types` for initial publish. Full typing needed for proper JSR docs.

### Files to Fix (by error count)

| File | Errors | Priority |
|------|--------|----------|
| `src/xlsx/xlsx.ts` | 89 | High (IO handler) |
| `src/doc/column.ts` | 48 | High (public API) |
| `src/stream/xlsx/worksheet-writer.ts` | 47 | Medium |
| `src/doc/row.ts` | 45 | High (public API) |
| `src/doc/worksheet.ts` | 43 | High (partially done) |
| `src/doc/table.ts` | 43 | Medium |
| `src/xlsx/xform/base-xform.ts` | 40 | Low (internal) |
| `src/doc/range.ts` | 40 | Medium |
| `src/doc/defined-names.ts` | 34 | Medium |
| `src/xlsx/xform/style/styles-xform.ts` | 27 | Low (internal) |
| `src/csv/csv.ts` | 15 | Medium |
| `src/utils/shared-strings.ts` | 7 | Low |
| `src/doc/data-validations.ts` | 7 | Medium |
| `src/utils/stream-buf.ts` | 2 | Low |

### Already Typed

- `src/doc/workbook.ts` ✅
- `src/doc/modelcontainer.ts` ✅
- `src/stream/xlsx/workbook-writer.ts` ✅
- `src/stream/xlsx/workbook-reader.ts` ✅

### Approach

1. Start with high-priority public API files (column, row, worksheet)
2. Add interface definitions for complex types
3. Use `unknown` for internal-only types to avoid cascading
4. Run `bunx jsr publish --dry-run` to verify progress

### Commands

```bash
# Check current error count
bunx jsr publish --dry-run 2>&1 | grep -c "error\[missing-explicit"

# List errors by file
bunx jsr publish --dry-run 2>&1 | grep -oP "src/[^:]+\.ts" | sort | uniq -c | sort -rn

# Publish with slow types (temporary)
bunx jsr publish --allow-slow-types
```

## Remove `any` Types

**Status:** 60 `any` type annotations remaining (down from 154 - 61% complete).

**Public API Files - Completed:**
- ✅ `src/xlsx/xlsx.ts` - Reduced from 42 to 2 (loop variables, acceptable)
- ✅ `src/stream/xlsx/worksheet-writer.ts` - Reduced from 25 to 0
- ✅ `src/doc/worksheet.ts` - Reduced from 11 to 0
- ✅ `src/doc/row.ts` - Reduced from 8 to 0
- ✅ `src/doc/cell.ts` - Reduced from 8 to 0
- ✅ `src/doc/column.ts` - Reduced from 5 to 1

**Remaining by Category:**
| File | Count | Type |
|------|-------|------|
| `src/doc/defined-names.ts` | 16 | Public API |
| `src/xlsx/xform/style/styles-xform.ts` | 12 | Internal |
| `src/xlsx/xform/base-xform.ts` | 7 | Internal |
| `src/doc/table.ts` | 7 | Public API |
| `src/csv/csv.ts` | 6 | Public API |
| Others | 12 | Low Priority |

**Next Steps:**
1. Create comprehensive type file for xforms and CSV
2. Tackle defined-names.ts and table.ts (public API)
3. Internal xform types can use relaxed typing (internal use)

### Files by `any` Count (Remaining)

| File | Count | Priority |
|------|-------|----------|
| `src/doc/defined-names.ts` | 16 | Medium |
| `src/xlsx/xform/style/styles-xform.ts` | 12 | Low (internal) |
| `src/doc/worksheet.ts` | 11 | High |
| `src/doc/row.ts` | 8 | High |
| `src/doc/cell.ts` | 8 | High |
| `src/xlsx/xform/base-xform.ts` | 7 | Low (internal) |
| `src/doc/table.ts` | 7 | Medium |
| `src/csv/csv.ts` | 6 | Medium |
| `src/doc/column.ts` | 5 | High |
| Others | 8 | Low |
| `src/doc/defined-names.ts` | 16 | Medium |
| `src/xlsx/xform/style/styles-xform.ts` | 12 | Low (internal) |
| `src/doc/worksheet.ts` | 11 | High |
| `src/doc/row.ts` | 8 | High |
| `src/doc/cell.ts` | 8 | High |
| `src/xlsx/xform/base-xform.ts` | 7 | Low (internal) |
| `src/doc/table.ts` | 7 | Medium |
| `src/csv/csv.ts` | 6 | Medium |
| `src/doc/column.ts` | 5 | High |
| `src/doc/range.ts` | 3 | Medium |
| `src/utils/stream-buf.ts` | 2 | Low |
| `src/doc/data-validations.ts` | 1 | Low |
| `src/csv/stream-converter.ts` | 1 | Low |

### Approach

1. **Create types file** (`src/types/index.ts`) for shared types (streams, common patterns)
2. **Leverage existing interfaces** (WorkbookModel, WorksheetModel, etc.)
3. **Fix high-impact files first** (xlsx.ts has 42 `any`s)
4. **Add Node.js stream types** from `readable-stream` package

### Type Categories to Define

- **Stream types**: `NodeJS.ReadableStream`, `NodeJS.WritableStream`, `PassThrough`
- **Model types**: Workbook/Worksheet/Cell models (some exist, need expansion)
- **Options types**: Read/write options, xform options
- **Zip types**: ZipStream.ZipWriter, JSZip entry types

### Commands

```bash
# Count any types
find src -name "*.ts" -type f | xargs rg ": any" | wc -l

# List by file
find src -name "*.ts" -type f | xargs rg ": any" | cut -d: -f1 | sort | uniq -c | sort -rn
```

## Other Tasks

- [ ] Update tests to use ExcelTS naming consistently
- [ ] Generate proper index.d.ts from source
- [ ] Add more JSDoc examples to public API
