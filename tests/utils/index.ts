// verquire is global from preload.ts

import _ from './under-dash.ts';
import tools from './tools.ts';
import testWorkbookReader from './test-workbook-reader.ts';
import _dataValidations from './test-data-validation-sheet.ts';
import _conditionalFormatting from './test-conditional-formatting-sheet.ts';
import _values from './test-values-sheet.ts';
import _splice from './test-spliced-sheet.ts';

// JSON imports
import viewsJson from './data/views.json';
import sheetValuesJson from './data/sheet-values.json';
import stylesJson from './data/styles.json';
import sheetPropertiesJson from './data/sheet-properties.json';
import pageSetupJson from './data/page-setup.json';
import conditionalFormattingJson from './data/conditional-formatting.json';
import headerFooterJson from './data/header-footer.json';

import Row from '../../src/doc/row.ts';
import Column from '../../src/doc/column.ts';
const testSheets = {
  dataValidations: _dataValidations,
  conditionalFormatting: _conditionalFormatting,
  values: _values,
  splice: _splice,
};

function getOptions(docType, options) {
  let result;
  switch (docType) {
    case 'xlsx':
      result = {
        sheetName: 'values',
        checkFormulas: true,
        checkMerges: true,
        checkStyles: true,
        checkBadAlignments: true,
        checkSheetProperties: true,
        dateAccuracy: 3,
        checkViews: true,
      };
      break;
    case 'csv':
      result = {
        sheetName: 'sheet1',
        checkFormulas: false,
        checkMerges: false,
        checkStyles: false,
        checkBadAlignments: false,
        checkSheetProperties: false,
        dateAccuracy: 1000,
        checkViews: false,
      };
      break;
    default:
      throw new Error(`Bad doc-type: ${docType}`);
  }
  return Object.assign(result, options);
}

export function createSheetMock() {
  return {
    _keys: {},
    _cells: {},
    rows: [],
    columns: [],
    properties: {
      outlineLevelCol: 0,
      outlineLevelRow: 0,
    },

    addColumn(colNumber, defn) {
      const newColumn = new Column(this, colNumber, defn);
      this.columns[colNumber - 1] = newColumn;
      return newColumn;
    },
    getColumn(colNumber) {
      let column = this.columns[colNumber - 1] || this._keys[colNumber];
      if (!column) {
        column = this.columns[colNumber - 1] = new Column(this, colNumber);
      }
      return column;
    },
    getRow(rowNumber) {
      let row = this.rows[rowNumber - 1];
      if (!row) {
        row = this.rows[rowNumber - 1] = new Row(this, rowNumber);
      }
      return row;
    },
    getCell(rowNumber, colNumber) {
      return this.getRow(rowNumber).getCell(colNumber);
    },
    getColumnKey(key) {
      return this._keys[key];
    },
    setColumnKey(key, value) {
      this._keys[key] = value;
    },
    deleteColumnKey(key) {
      delete this._keys[key];
    },
    eachColumnKey(f) {
      _.each(this._keys, f);
    },
    eachRow(opt, f) {
      if (!f) {
        f = opt;
        opt = {};
      }
      if (opt && opt.includeEmpty) {
        const n = this.rows.length;
        for (let i = 1; i <= n; i++) {
          f(this.getRow(i), i);
        }
      } else {
        this.rows.forEach((r, i) => {
          if (r) {
            f(r, i + 1);
          }
        });
      }
    },
  };
}

export default {
  views: tools.fix(viewsJson),
  testValues: tools.fix(sheetValuesJson),
  styles: tools.fix(stylesJson),
  properties: tools.fix(sheetPropertiesJson),
  pageSetup: tools.fix(pageSetupJson),
  conditionalFormatting: tools.fix(conditionalFormattingJson),
  headerFooter: tools.fix(headerFooterJson),

  createTestBook(workbook, docType, sheets) {
    const options = getOptions(docType);
    sheets = sheets || ['values'];

    workbook.views = [
      {x: 1, y: 2, width: 10000, height: 20000, firstSheet: 0, activeTab: 0},
    ];

    sheets.forEach(sheet => {
      const testSheet = _.get(testSheets, sheet);
      testSheet.addSheet(workbook, options);
    });

    return workbook;
  },

  checkTestBook(workbook, docType, sheets, options) {
    options = getOptions(docType, options);
    sheets = sheets || ['values'];

    expect(workbook).toBeDefined();

    if (options.checkViews) {
      expect(workbook.views).toEqual([
        {
          x: 1,
          y: 2,
          width: 10000,
          height: 20000,
          firstSheet: 0,
          activeTab: 0,
          visibility: 'visible',
        },
      ]);
    }

    sheets.forEach(sheet => {
      const testSheet = _.get(testSheets, sheet);
      testSheet.checkSheet(workbook, options);
    });
  },

  checkTestBookReader: testWorkbookReader.checkBook,

  createSheetMock,
};
