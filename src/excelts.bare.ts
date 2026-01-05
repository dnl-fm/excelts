/**
 * Browser entrypoint without polyfills.
 */
import Enums from './doc/enums.ts';
import _Workbook from './doc/workbook.ts';

const ExcelJS = {
  Workbook: _Workbook,
};

Object.keys(Enums).forEach(key => {
  ExcelJS[key] = Enums[key];
});

const globalScope = typeof self !== 'undefined' ? self : typeof window !== 'undefined' ? window : undefined;
if (globalScope) {
  globalScope.ExcelJS = ExcelJS;
}

export default ExcelJS;