/**
 * Node-style entrypoint for ExcelJS APIs and stream helpers.
 */
import _Workbook from './doc/workbook.ts';
import _ModelContainer from './doc/modelcontainer.ts';
import _WorkbookWriter from './stream/xlsx/workbook-writer.ts';
import _WorkbookReader from './stream/xlsx/workbook-reader.ts';
import * as _enums from './doc/enums.ts';

const ExcelJS = {
  Workbook: _Workbook,
  ModelContainer: _ModelContainer,
  stream: {
    xlsx: {
      WorkbookWriter: _WorkbookWriter,
      WorkbookReader: _WorkbookReader,
    },
  },
};

Object.assign(ExcelJS, _enums);

export default ExcelJS;