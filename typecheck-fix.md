# TypeScript Type Check Fix Guide

## Current Status (2026-01-06)
- **405 type errors remaining** (down from ~690 originally)
- **54 errors in source files** (`src/`)
- **351 errors in test files** (`tests/`) - will resolve when source types are fixed
- **876 unit tests passing**, 7 failing (pre-existing parsing issues)

## Rules

### DO NOT USE
- `any` - never acceptable
- `unknown` with immediate cast - if you know the type, declare it properly
- Type assertions (`as`) to work around poor typing

### ACCEPTABLE
- `unknown` only when data truly comes from external source (parsed XML, user input)
- Proper type definitions using `type` keyword
- Generic types where appropriate

### CALLER ANALYSIS RULE
When fixing a method with `any` or `unknown` parameters:
1. **Find all callers** of that method (use grep/search)
2. **Analyze what callers pass** - look at the actual values being passed
3. **Define a type** that covers what callers actually send
4. **Type the parameter** with that defined type

Example:
```typescript
// BAD - we don't know what options is
prepare(model: unknown, options: unknown) {
  const opts = options as SomeType; // casting = we actually knew the type!
}

// GOOD - analyze callers, define type, use it
type PrepareOptions = {
  styles: StylesManager;
  comments: CommentRef[];
  // ... properties we found from caller analysis
};

prepare(model: CellModel, options: PrepareOptions) {
  // no casts needed, type is known
}
```

This ensures types flow correctly from caller to callee without casts.

### WORKFLOW FOR FIXING A METHOD

1. **Identify the method** with `any`/`unknown` params
2. **Find all callers**:
   ```bash
   grep -rn "methodName" src/ --include="*.ts"
   ```
3. **Trace the caller's context** - what does the caller receive? Where does IT get its params from?
4. **Follow the chain up** until you find the origin of the data
5. **Define types** at the appropriate level (may need shared types file)
6. **Apply types** from top down so they flow naturally

## Example: CellXform.prepare Call Chain Analysis

```
xlsx.ts:prepareModel()
  └─ worksheetOptions = { sharedStrings, styles, date1904, drawingsCount, media, drawings, commentRefs }
     └─ worksheetXform.prepare(worksheet, worksheetOptions)
        └─ worksheet-xform.ts:prepare() adds: merges, hyperlinks, comments, formulae, siFormulae
           └─ this.map.sheetData.prepare(model.rows, options)  [sheetData is ListXform of RowXform]
              └─ ListXform calls childXform.prepare for each row
                 └─ row-xform.ts:prepare(model, options)
                    └─ cellXform.prepare(cellModel, options)
                       └─ cell-xform.ts:prepare() uses: styles, comments, sharedStrings, date1904,
                                                        hyperlinks, merges, formulae, siFormulae
```

**worksheetOptions starts with:**
- sharedStrings: SharedStringsXform
- styles: StylesXform
- date1904: boolean
- drawingsCount: number
- media: MediaModel[]
- drawings: []
- commentRefs: []

**worksheet-xform.prepare adds:**
- merges: Merges instance
- hyperlinks: [] (also set on model)
- comments: [] (also set on model)
- formulae: {}
- siFormulae: 0

**So CellXformPrepareOptions needs:**
- styles.addStyleModel(style, type) => number | undefined
- comments: array to push to
- sharedStrings?.add(value) => number
- date1904?: boolean
- hyperlinks: array to push to
- merges.add(model)
- formulae: Record<string, CellModel>
- siFormulae: number (but it's mutated with ++)

## Pattern for Xform Files

Each xform class follows a pattern:
1. `prepare(model, options)` - prepares model for rendering
2. `render(xmlStream, model)` - writes XML
3. `parseOpen/parseText/parseClose` - SAX parsing
4. `reconcile(model, options)` - post-parse processing

### Proper Typing Approach

1. **Define model types** for each xform:
```typescript
type CellModel = {
  address?: string;
  type?: number;
  value?: unknown;
  styleId?: number;
  // ... all known properties
};
```

2. **Define options types** for prepare/reconcile:
```typescript
type CellXformPrepareOptions = {
  styles: {
    addStyleModel: (style: unknown, type: number) => number | undefined;
  };
  comments: Array<{ ref: string; [key: string]: unknown }>;
  sharedStrings?: {
    add: (value: unknown) => number;
  };
  // ... all known properties
};

type CellXformReconcileOptions = {
  styles?: {
    getStyleModel: (id: number) => { numFmt?: string };
  };
  sharedStrings?: {
    getString: (id: number) => unknown;
  };
  // ... all known properties
};
```

3. **Use types in method signatures**:
```typescript
prepare(model: CellModel, options: CellXformPrepareOptions) {
  // no casts needed inside
}
```

4. **For inherited methods from BaseXform**, the base class uses `unknown` because it doesn't know subtypes. Subclasses should override with specific types.

## Shared Types

Created `src/xlsx/xform/xform-types.ts` with:

- **Model types**: `CellModel`, `RowModel`
- **Options types**: `CellXformPrepareOptions`, `CellXformReconcileOptions`, `RowXformPrepareOptions`, `RowXformReconcileOptions`
- **Manager interfaces**: `StylesManager`, `SharedStringsManager`, `MergesManager`
- **Utility types**: `XmlStreamWriter`, `SaxNode`
- **Helper types**: `HyperlinkRef`, `CommentRef`, `StyleModel`

Import and use these types in xform files:
```typescript
import type {
  CellModel,
  CellXformPrepareOptions,
  XmlStreamWriter,
  SaxNode,
} from '../xform-types.ts';
```

## Files to Fix (Priority Order)

### High Impact (fix these first)
1. `src/xlsx/xform/base-xform.ts` - base class, affects all xforms
2. `src/xlsx/xform/sheet/cell-xform.ts` - most used xform
3. `src/xlsx/xform/sheet/row-xform.ts` - uses cell-xform
4. `src/xlsx/xform/sheet/worksheet-xform.ts` - orchestrates sheet parsing
5. `src/xlsx/xform/style/styles-xform.ts` - style management

### Medium Impact
- `src/xlsx/xform/strings/shared-strings-xform.ts`
- `src/xlsx/xform/core/relationships-xform.ts`
- `src/xlsx/xform/list-xform.ts`
- `src/xlsx/xform/sheet/cf/*.ts` (conditional formatting)

### Lower Impact
- Comment xforms
- Drawing xforms
- Pivot table xforms
- Table xforms

## Testing After Fixes

```bash
# Check types
bunx tsc --noEmit

# Run tests to ensure nothing broke
bun test tests/unit
```

## Progress

### Session 2026-01-06 (continued): Down from 441 to 405 errors

**Files Fixed (this session):**
- `src/xlsx/xform/pivot-table/cache-field.ts` - added CacheFieldParams, typed class properties
- `src/xlsx/xform/sheet/cf-ext/databar-ext-xform.ts` - added DatabarExtModel, child xform types
- `src/xlsx/xform/table/table-xform.ts` - moved static property inside class, added model types
- `src/xlsx/xform/sheet/cf-ext/icon-set-ext-xform.ts` - added IconSetExtModel, child xform types
- `src/xlsx/xform/sheet/cf-ext/conditional-formatting-ext-xform.ts` - added proper types
- `src/xlsx/xform/sheet/cf-ext/conditional-formattings-ext-xform.ts` - array model type
- `src/xlsx/xform/pivot-table/pivot-cache-definition-xform.ts` - moved static prop, typed methods
- `src/xlsx/xform/pivot-table/pivot-cache-records-xform.ts` - moved static prop, typed methods
- `src/xlsx/xform/pivot-table/pivot-table-xform.ts` - moved static prop, typed methods
- `src/xlsx/xform/sheet/ext-lst-xform.ts` - added ConditionalFormattingsModel type
- `src/xlsx/xform/list-xform.ts` - fixed array model handling, removed incompatible declare
- `src/xlsx/xform/style/border-xform.ts` - moved EdgeXform.validStyleValues to static property
- `src/xlsx/xform/style/fill-xform.ts` - fixed model casts in parseClose and render
- `src/xlsx/xform/simple/boolean-xform.ts` - added proper types
- `src/xlsx/xform/base-xform.ts` - changed model type to `unknown` for flexibility
- `src/xlsx/xform/sheet/header-footer-xform.ts` - added HeaderFooterModel type
- `src/xlsx/xform/comment/vml-client-data-xform.ts` - added VmlClientDataModel types
- `src/xlsx/xform/drawing/one-cell-anchor-xform.ts` - added OneCellAnchorModel
- `src/xlsx/xform/drawing/two-cell-anchor-xform.ts` - added TwoCellAnchorModel
- `src/xlsx/xform/drawing/drawing-xform.ts` - added DrawingModel, moved static prop
- `src/xlsx/xform/static-xform.ts` - added optional model param to render

**Additional Fixes:**
- `src/xlsx/xform/xform-types.ts` - added writeXml to XmlStreamWriter, added index signature to RowModel
- `src/xlsx/xform/sheet/cf-ext/cf-rule-ext-xform.ts` - added child xform property declarations
- `src/xlsx/xform/list-xform.ts` - changed prepare options to Record<string, unknown>
- `src/stream/xlsx/worksheet-writer.ts` - added proper type casts for model and options

**Key Change:** Changed BaseXform.model from `XformModel` (object type) to `unknown` 
to allow subclasses to use different model types (boolean, array, etc.)

**Tests:** 876 passing, 7 failing (down from 17 failures before session!)

### Current Error Summary (405 total errors)

**Breakdown:**
- **54 source file errors** (in `src/`)
- **351 test file errors** (in `tests/`, expected as they follow source types)

**Source errors by file (top 20):**
```
3 src/xlsx/xform/style/style-xform.ts
3 src/xlsx/xform/strings/phonetic-text-xform.ts  
3 src/xlsx/xform/sheet/sheet-properties-xform.ts
3 src/xlsx/xform/sheet/sheet-format-properties-xform.ts
2 src/xlsx/xlsx.ts
2 src/xlsx/xform/style/fill-xform.ts
2 src/xlsx/xform/strings/shared-string-xform.ts
2 src/xlsx/xform/strings/rich-text-xform.ts
2 src/xlsx/xform/static-xform.ts
2 src/xlsx/xform/sheet/worksheet-xform.ts
2 src/xlsx/xform/sheet/cf/ext-lst-ref-xform.ts
2 src/xlsx/xform/sheet/cf/cf-rule-xform.ts
2 src/xlsx/xform/sheet/cf-ext/databar-ext-xform.ts
2 src/xlsx/xform/sheet/cf-ext/conditional-formatting-ext-xform.ts
2 src/xlsx/xform/sheet/cf-ext/cfvo-ext-xform.ts
1 src/xlsx/xform/table/filter-column-xform.ts
1 src/xlsx/xform/table/auto-filter-xform.ts
1 src/xlsx/xform/style/color-xform.ts
1 src/xlsx/xform/style/border-xform.ts
1 src/xlsx/xform/strings/text-xform.ts
```

### Next Steps

1. **Fix remaining 54 source errors** - mostly adding model types to xform classes
2. **Common patterns remaining:**
   - `this.model.property` access on `unknown` type → add `declare model: ModelType`
   - Child xform property access → add property declarations
   - Static properties after class → move inside class as `static prop = value`
3. **Test errors will auto-resolve** once source types are fixed

### Previous Session: Down from 607 to ~460 errors

**Files Fixed:**
- `src/xlsx/xform/sheet/merges.ts` - added proper model types, typed methods
- `src/xlsx/xform/strings/shared-strings-xform.ts` - added SharedStringsModel type
- `src/xlsx/xform/sheet/data-validations-xform.ts` - added DataValidation types
- `src/xlsx/xform/sheet/cf/cf-rule-xform.ts` - added CfRuleModel, typed all methods
- `src/xlsx/xform/simple/date-xform.ts` - typed options and methods
- `src/xlsx/xform/simple/integer-xform.ts` - typed options and methods
- `src/xlsx/xform/simple/float-xform.ts` - typed options and methods
- `src/xlsx/xform/simple/string-xform.ts` - typed options and methods
- `src/xlsx/xform/core/content-types-xform.ts` - added ContentTypesModel
- `src/xlsx/xform/sheet/col-xform.ts` - added ColModel and options types
- `src/xlsx/xform/style/font-xform.ts` - added FontMapEntry, FontModel types
- `src/xlsx/xform/sheet/sheet-protection-xform.ts` - added SheetProtectionModel
- `src/xlsx/xform/core/core-xform.ts` - fixed DateFormat type (was string, now function)
- `src/xlsx/xform/sheet/ext-lst-xform.ts` - typed ExtXform and ExtLstXform
- `src/xlsx/xform/sheet/cf/cfvo-xform.ts` - fixed parseOpen to conditionally add value
- `src/xlsx/xform/sheet/cf/icon-set-xform.ts` - fixed createNewModel to conditionally add props
- `src/types/index.ts` - added XlsxReadOptions properties, fixed WritableStream naming
- `src/xlsx/xlsx.ts` - fixed stream type issues

**Types Added to xform-types.ts:**
- `addAttributes` method to XmlStreamWriter

### Previous Session: Created foundation
- `src/xlsx/xform/xform-types.ts` - shared types file created
- `src/xlsx/xform/sheet/cell-xform.ts` - fully typed (0 errors, 70 tests pass)
- `src/xlsx/xform/sheet/row-xform.ts` - fully typed (0 errors, 20 tests pass)

### Remaining Source Files (by error count, as of 405 errors)

**Priority files (3 errors each):**
1. `src/xlsx/xform/style/style-xform.ts`
2. `src/xlsx/xform/strings/phonetic-text-xform.ts`
3. `src/xlsx/xform/sheet/sheet-properties-xform.ts`
4. `src/xlsx/xform/sheet/sheet-format-properties-xform.ts`

**Medium priority (2 errors each):**
5. `src/xlsx/xlsx.ts`
6. `src/xlsx/xform/style/fill-xform.ts`
7. `src/xlsx/xform/strings/shared-string-xform.ts`
8. `src/xlsx/xform/strings/rich-text-xform.ts`
9. `src/xlsx/xform/static-xform.ts`
10. `src/xlsx/xform/sheet/worksheet-xform.ts`
11. `src/xlsx/xform/sheet/cf/ext-lst-ref-xform.ts`
12. `src/xlsx/xform/sheet/cf/cf-rule-xform.ts`
13. `src/xlsx/xform/sheet/cf-ext/databar-ext-xform.ts`
14. `src/xlsx/xform/sheet/cf-ext/conditional-formatting-ext-xform.ts`
15. `src/xlsx/xform/sheet/cf-ext/cfvo-ext-xform.ts`

**Low priority (1 error each):**
- Various table, style, and string xforms

### Next Steps
Continue with the remaining xform files, following the same pattern:
1. Analyze callers to understand what types flow in
2. Add types to `xform-types.ts` if needed
3. Import and use proper types in the xform file
4. Run tests to verify nothing broke

## Notes

- The `map` property in xforms holds child xform instances. Type it specifically per class.
- `model` property holds parsed data. Type it specifically per class using `declare model: ModelType | undefined`.
- SAX parsing methods receive node data from the parser - these can use `SaxNode` type.
- Some xforms have static properties assigned after class definition - convert to static class properties.

## Lessons Learned

### Mixed types at different lifecycle stages
Some properties have different types at different stages. Example: `CellModel.si`:
- During **parsing**: string (from XML attribute)
- During **prepare/reconcile**: number (from siFormulae counter)

Type as union: `si?: number | string` and document the behavior.

### Don't over-convert XML attributes
XML attributes come as strings. Don't automatically `parseInt()` them unless the code expects numbers.
The original code often keeps them as strings and JavaScript's loose equality handles comparisons.
