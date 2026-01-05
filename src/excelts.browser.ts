/**
 * Browser entrypoint that exposes ExcelTS on the global scope.
 */
import Enums from './doc/enums.ts';
import _Workbook from './doc/workbook.ts';

const ExcelTS = {
  Workbook: _Workbook,
};

Object.keys(Enums).forEach(key => {
  ExcelTS[key] = Enums[key];
});

const globalScope = typeof self !== 'undefined' ? self : typeof window !== 'undefined' ? window : undefined;
if (globalScope) {
  globalScope.ExcelTS = ExcelTS;
}

export default ExcelTS;
