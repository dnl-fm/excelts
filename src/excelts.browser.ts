/**
 * Browser entrypoint that exposes ExcelTS on the global scope.
 * No Node.js polyfills required - uses native browser APIs.
 */

import Enums from './doc/enums.ts';
import _Workbook from './doc/workbook.ts';

declare global {
  interface Window {
    ExcelTS: typeof ExcelTS;
  }
}

const ExcelTS = {
  Workbook: _Workbook,
};

Object.keys(Enums).forEach(key => {
  ExcelTS[key as keyof typeof Enums] = Enums[key as keyof typeof Enums];
});

const globalScope = typeof self !== 'undefined' ? self : typeof window !== 'undefined' ? window : undefined;
if (globalScope) {
  (globalScope as unknown as Window).ExcelTS = ExcelTS;
}

export default ExcelTS;
