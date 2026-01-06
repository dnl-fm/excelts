/**
 * ExcelTS - Read and write Excel XLSX and CSV files.
 *
 * @example Basic usage
 * ```ts
 * import ExcelTS from '@dnl-fm/excelts';
 *
 * // Create a workbook
 * const workbook = new ExcelTS.Workbook();
 * const sheet = workbook.addWorksheet('Sheet1');
 *
 * sheet.columns = [
 *   { header: 'Name', key: 'name', width: 20 },
 *   { header: 'Age', key: 'age', width: 10 },
 * ];
 *
 * sheet.addRow({ name: 'Alice', age: 30 });
 * sheet.addRow({ name: 'Bob', age: 25 });
 *
 * await workbook.xlsx.writeFile('output.xlsx');
 * ```
 *
 * @example Streaming large files
 * ```ts
 * import ExcelTS from '@dnl-fm/excelts';
 *
 * const workbook = new ExcelTS.stream.xlsx.WorkbookWriter({
 *   filename: 'large.xlsx',
 * });
 *
 * const sheet = workbook.addWorksheet('Data');
 * for (let i = 0; i < 1000000; i++) {
 *   sheet.addRow([i, `Row ${i}`]).commit();
 * }
 *
 * await sheet.commit();
 * await workbook.commit();
 * ```
 *
 * @module
 */
import _Workbook from './doc/workbook.ts';
import _ModelContainer from './doc/modelcontainer.ts';
import _WorkbookWriter from './stream/xlsx/workbook-writer.ts';
import _WorkbookReader from './stream/xlsx/workbook-reader.ts';
import { ValueType, FormulaType, RelationshipType, DocumentType, ReadingOrder, ErrorValue } from './doc/enums.ts';

// Re-export classes for named imports
/** Main Workbook class for creating and manipulating Excel files. */
export const Workbook = _Workbook;
/** Container for workbook model data. */
export const ModelContainer = _ModelContainer;
/** Streaming writer for large XLSX files. */
export const WorkbookWriter = _WorkbookWriter;
/** Streaming reader for large XLSX files. */
export const WorkbookReader = _WorkbookReader;

// Re-export enums
export { ValueType, FormulaType, RelationshipType, DocumentType, ReadingOrder, ErrorValue };

/**
 * ExcelTS namespace object for convenient access to all APIs.
 *
 * @example
 * ```ts
 * import ExcelTS from '@dnl-fm/excelts';
 * const workbook = new ExcelTS.Workbook();
 * ```
 */
const ExcelTS = {
  Workbook: _Workbook,
  ModelContainer: _ModelContainer,
  stream: {
    xlsx: {
      WorkbookWriter: _WorkbookWriter,
      WorkbookReader: _WorkbookReader,
    },
  },
  ValueType,
  FormulaType,
  RelationshipType,
  DocumentType,
  ReadingOrder,
  ErrorValue,
};

export default ExcelTS;
