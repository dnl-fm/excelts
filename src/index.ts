/**
 * Main entrypoint for ExcelTS APIs and stream helpers.
 */
import _Workbook from './doc/workbook.ts';
import _ModelContainer from './doc/modelcontainer.ts';
import _WorkbookWriter from './stream/xlsx/workbook-writer.ts';
import _WorkbookReader from './stream/xlsx/workbook-reader.ts';
import { ValueType, FormulaType, RelationshipType, DocumentType, ReadingOrder, ErrorValue } from './doc/enums.ts';

const ExcelTS = {
  Workbook: _Workbook,
  ModelContainer: _ModelContainer,
  stream: {
    xlsx: {
      WorkbookWriter: _WorkbookWriter,
      WorkbookReader: _WorkbookReader,
    },
  },
  // Enums
  ValueType,
  FormulaType,
  RelationshipType,
  DocumentType,
  ReadingOrder,
  ErrorValue,
};

export default ExcelTS;
