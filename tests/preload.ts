/**
 * Preload script for Bun tests
 * Sets up global verquire helper using dynamic imports
 */
import path from 'path';

const projectRoot = path.resolve(import.meta.dir, '..');

// Cache for loaded modules
const libs: Record<string, unknown> = {};

// Pre-load exceljs
const exceljs = await import('../src/excelts.nodejs.js');
libs.exceljs = exceljs.default;

// Pre-load commonly used modules dynamically
const preloadModules = [
  // doc
  'doc/workbook',
  'doc/worksheet',
  'doc/cell',
  'doc/row',
  'doc/column',
  'doc/anchor',
  'doc/enums',
  'doc/defined-names',
  'doc/range',
  'doc/image',
  'doc/note',
  'doc/modelcontainer',
  // utils
  'utils/col-cache',
  'utils/under-dash',
  'utils/copy-style',
  'utils/stream-buf',
  'utils/string-buf',
  'utils/shared-formula',
  'utils/utils',
  'utils/xml-stream',
  'utils/parse-sax',
  'utils/shared-strings',
  'utils/cell-matrix',
  // xlsx/xform
  'xlsx/xform/base-xform',
  'xlsx/xform/composite-xform',
  'xlsx/xform/static-xform',
  'xlsx/xform/list-xform',
  // xlsx/xform/simple
  'xlsx/xform/simple/boolean-xform',
  'xlsx/xform/simple/integer-xform',
  'xlsx/xform/simple/string-xform',
  'xlsx/xform/simple/float-xform',
  'xlsx/xform/simple/date-xform',
  // xlsx/xform/style
  'xlsx/xform/style/styles-xform',
  'xlsx/xform/style/style-xform',
  'xlsx/xform/style/font-xform',
  'xlsx/xform/style/fill-xform',
  'xlsx/xform/style/border-xform',
  'xlsx/xform/style/numfmt-xform',
  'xlsx/xform/style/alignment-xform',
  'xlsx/xform/style/protection-xform',
  'xlsx/xform/style/color-xform',
  'xlsx/xform/style/underline-xform',
  'xlsx/xform/style/dxf-xform',
  // xlsx/xform/strings
  'xlsx/xform/strings/shared-strings-xform',
  'xlsx/xform/strings/shared-string-xform',
  'xlsx/xform/strings/rich-text-xform',
  'xlsx/xform/strings/text-xform',
  'xlsx/xform/strings/phonetic-text-xform',
  // xlsx/xform/core
  'xlsx/xform/core/core-xform',
  'xlsx/xform/core/app-xform',
  'xlsx/xform/core/relationships-xform',
  'xlsx/xform/core/relationship-xform',
  'xlsx/xform/core/content-types-xform',
  'xlsx/xform/core/app-heading-pairs-xform',
  'xlsx/xform/core/app-titles-of-parts-xform',
  // xlsx/xform/book
  'xlsx/xform/book/workbook-xform',
  'xlsx/xform/book/sheet-xform',
  'xlsx/xform/book/defined-name-xform',
  'xlsx/xform/book/workbook-view-xform',
  'xlsx/xform/book/workbook-properties-xform',
  'xlsx/xform/book/workbook-calc-properties-xform',
  // xlsx/xform/sheet
  'xlsx/xform/sheet/worksheet-xform',
  'xlsx/xform/sheet/row-xform',
  'xlsx/xform/sheet/cell-xform',
  'xlsx/xform/sheet/col-xform',
  'xlsx/xform/sheet/dimension-xform',
  'xlsx/xform/sheet/hyperlink-xform',
  'xlsx/xform/sheet/merge-cell-xform',
  'xlsx/xform/sheet/merges',
  'xlsx/xform/sheet/data-validations-xform',
  'xlsx/xform/sheet/sheet-view-xform',
  'xlsx/xform/sheet/sheet-format-properties-xform',
  'xlsx/xform/sheet/sheet-properties-xform',
  'xlsx/xform/sheet/sheet-protection-xform',
  'xlsx/xform/sheet/page-margins-xform',
  'xlsx/xform/sheet/page-setup-xform',
  'xlsx/xform/sheet/page-setup-properties-xform',
  'xlsx/xform/sheet/page-breaks-xform',
  'xlsx/xform/sheet/row-breaks-xform',
  'xlsx/xform/sheet/print-options-xform',
  'xlsx/xform/sheet/header-footer-xform',
  'xlsx/xform/sheet/auto-filter-xform',
  'xlsx/xform/sheet/picture-xform',
  'xlsx/xform/sheet/drawing-xform',
  'xlsx/xform/sheet/outline-properties-xform',
  'xlsx/xform/sheet/ext-lst-xform',
  'xlsx/xform/sheet/table-part-xform',
  // xlsx/xform/sheet/cf
  'xlsx/xform/sheet/cf/cf-rule-xform',
  'xlsx/xform/sheet/cf/cfvo-xform',
  'xlsx/xform/sheet/cf/color-scale-xform',
  'xlsx/xform/sheet/cf/databar-xform',
  'xlsx/xform/sheet/cf/ext-lst-ref-xform',
  'xlsx/xform/sheet/cf/formula-xform',
  'xlsx/xform/sheet/cf/icon-set-xform',
  'xlsx/xform/sheet/cf/conditional-formatting-xform',
  'xlsx/xform/sheet/cf/conditional-formattings-xform',
  // xlsx/xform/sheet/cf-ext
  'xlsx/xform/sheet/cf-ext/cf-rule-ext-xform',
  'xlsx/xform/sheet/cf-ext/cfvo-ext-xform',
  'xlsx/xform/sheet/cf-ext/databar-ext-xform',
  'xlsx/xform/sheet/cf-ext/icon-set-ext-xform',
  'xlsx/xform/sheet/cf-ext/cf-icon-ext-xform',
  'xlsx/xform/sheet/cf-ext/f-ext-xform',
  'xlsx/xform/sheet/cf-ext/sqref-ext-xform',
  'xlsx/xform/sheet/cf-ext/conditional-formatting-ext-xform',
  'xlsx/xform/sheet/cf-ext/conditional-formattings-ext-xform',
  // xlsx/xform/table
  'xlsx/xform/table/table-xform',
  'xlsx/xform/table/table-column-xform',
  'xlsx/xform/table/table-style-info-xform',
  'xlsx/xform/table/auto-filter-xform',
  'xlsx/xform/table/filter-column-xform',
  'xlsx/xform/table/filter-xform',
  'xlsx/xform/table/custom-filter-xform',
  // xlsx/xform/drawing
  'xlsx/xform/drawing/drawing-xform',
  'xlsx/xform/drawing/two-cell-anchor-xform',
  'xlsx/xform/drawing/one-cell-anchor-xform',
  'xlsx/xform/drawing/pic-xform',
  'xlsx/xform/drawing/blip-xform',
  'xlsx/xform/drawing/blip-fill-xform',
  'xlsx/xform/drawing/cell-position-xform',
  // xlsx
  'xlsx/xlsx',
  // stream
  'stream/xlsx/workbook-writer',
  'stream/xlsx/workbook-reader',
  'stream/xlsx/worksheet-writer',
  'stream/xlsx/worksheet-reader',
];

for (const mod of preloadModules) {
  try {
    const imported = await import(path.join(projectRoot, 'src', mod + '.js'));
    libs[mod] = imported.default;
  } catch {
    // Module may not exist, skip
    console.warn(`Warning: Could not preload module: ${mod}`);
  }
}

globalThis.verquire = (modulePath: string) => {
  if (!libs[modulePath]) {
    throw new Error(
      `Module "${modulePath}" not pre-loaded. Add it to tests/preload.ts or import directly in your test.`
    );
  }
  return libs[modulePath];
};

// Extend global for TypeScript
declare global {
  var verquire: (modulePath: string) => any;
}
