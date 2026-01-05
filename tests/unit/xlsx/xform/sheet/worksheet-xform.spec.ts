import fs from 'fs';
import testXformHelper from '../test-xform-helper.ts';

// JSON imports - Sheet 1
import sheet10Json from './data/sheet.1.0.json';
import sheet11Json from './data/sheet.1.1.json';
import sheet13Json from './data/sheet.1.3.json';
import sheet14Json from './data/sheet.1.4.json';

// JSON imports - Sheet 2
import sheet20Json from './data/sheet.2.0.json';
import sheet21Json from './data/sheet.2.1.json';

// JSON imports - Sheet 3
import sheet31Json from './data/sheet.3.1.json';

// JSON imports - Sheet 4
import sheet40Json from './data/sheet.4.0.json';

// JSON imports - Sheet 5
import sheet50Json from './data/sheet.5.0.json';
import sheet51Json from './data/sheet.5.1.json';
import sheet53Json from './data/sheet.5.3.json';
import sheet54Json from './data/sheet.5.4.json';

// JSON imports - Sheet 6
import sheet61Json from './data/sheet.6.1.json';
import sheet63Json from './data/sheet.6.3.json';

// JSON imports - Sheet 7
import sheet70Json from './data/sheet.7.0.json';
import sheet71Json from './data/sheet.7.1.json';

import Enums from '../../../../../src/doc/enums.ts';
import XmlStream from '../../../../../src/utils/xml-stream.ts';
import WorksheetXform from '../../../../../src/xlsx/xform/sheet/worksheet-xform.ts';
import SharedStringsXform from '../../../../../src/xlsx/xform/strings/shared-strings-xform.ts';
import StylesXform from '../../../../../src/xlsx/xform/style/styles-xform.ts';
const fakeStyles = {
  addStyleModel(style, cellType) {
    if (cellType === Enums.ValueType.Date) {
      return 1;
    }
    if (style && style.font) {
      return 2;
    }
    return 0;
  },
  getStyleModel(id) {
    switch (id) {
      case 1:
        return {numFmt: 'mm-dd-yy'};
      case 2:
        return {
          font: {
            underline: true,
            size: 11,
            color: {theme: 10},
            name: 'Calibri',
            family: 2,
            scheme: 'minor',
          },
        };
      default:
        return null;
    }
  },
};

const fakeHyperlinkMap = {
  B6: 'https://www.npmjs.com/package/exceljs',
};

function fixDate(model) {
  // Clone the model to avoid mutating the imported JSON
  const clone = JSON.parse(JSON.stringify(model));
  clone.rows[3].cells[1].value = new Date(clone.rows[3].cells[1].value);
  return clone;
}

const expectations = [
  {
    title: 'Sheet 1',
    create: () => new WorksheetXform(),
    initialModel: fixDate(sheet10Json),
    preparedModel: fixDate(sheet11Json),
    xml: fs.readFileSync(`${__dirname}/data/sheet.1.2.xml`).toString(),
    parsedModel: sheet13Json,
    reconciledModel: fixDate(sheet14Json),
    tests: ['prepare', 'render', 'parse'],
    options: {
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles,
      formulae: {},
      siFormulae: 0,
    },
  },
  {
    title: 'Sheet 2 - Data Validations',
    create: () => new WorksheetXform(),
    initialModel: sheet20Json,
    preparedModel: sheet21Json,
    xml: fs.readFileSync(`${__dirname}/data/sheet.2.2.xml`).toString(),
    tests: ['prepare', 'render'],
    options: {
      styles: new StylesXform(true),
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
      formulae: {},
      siFormulae: 0,
    },
  },
  {
    title: 'Sheet 3 - Empty Sheet',
    create: () => new WorksheetXform(),
    preparedModel: sheet31Json,
    xml: fs.readFileSync(`${__dirname}/data/sheet.3.2.xml`).toString(),
    tests: ['render'],
    options: {
      styles: new StylesXform(true),
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
    },
  },
  {
    title: 'Sheet 5 - Shared Formulas',
    create: () => new WorksheetXform(),
    initialModel: sheet50Json,
    preparedModel: sheet51Json,
    xml: fs.readFileSync(`${__dirname}/data/sheet.5.2.xml`).toString(),
    parsedModel: sheet53Json,
    reconciledModel: sheet54Json,
    tests: ['prepare-render', 'parse'],
    options: {
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles,
      formulae: {},
      siFormulae: 0,
    },
  },
  {
    title: 'Sheet 6 - AutoFilter',
    create: () => new WorksheetXform(),
    preparedModel: sheet61Json,
    xml: fs.readFileSync(`${__dirname}/data/sheet.6.2.xml`).toString(),
    parsedModel: sheet63Json,
    tests: ['render', 'parse'],
    options: {
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles,
      formulae: {},
      siFormulae: 0,
    },
  },
  {
    title: 'Sheet 7 - Row Breaks',
    create: () => new WorksheetXform(),
    initialModel: sheet70Json,
    preparedModel: sheet71Json,
    xml: fs.readFileSync(`${__dirname}/data/sheet.7.2.xml`).toString(),
    tests: ['prepare', 'render'],
    options: {
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles,
      formulae: {},
      siFormulae: 0,
    },
  },
];

describe('WorksheetXform', () => {
  testXformHelper(expectations);

  it('hyperlinks must be after dataValidations', () => {
    const xform = new WorksheetXform();
    const model = JSON.parse(JSON.stringify(sheet40Json));
    const xmlStream = new XmlStream();
    const options = {
      styles: new StylesXform(true),
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
    };
    xform.prepare(model, options);
    xform.render(xmlStream, model);

    const {xml} = xmlStream;
    const iHyperlinks = xml.indexOf('hyperlinks');
    const iDataValidations = xml.indexOf('dataValidations');
    expect(iHyperlinks).not.toBe(-1);
    expect(iDataValidations).not.toBe(-1);
    expect(iHyperlinks).toBeGreaterThan(iDataValidations);
  });

  it('conditionalFormattings must be before dataValidations', () => {
    const xform = new WorksheetXform();
    const model = JSON.parse(JSON.stringify(sheet40Json));
    const xmlStream = new XmlStream();
    const options = {
      styles: new StylesXform(true),
      hyperlinks: [],
    };
    xform.prepare(model, options);
    xform.render(xmlStream, model);

    const {xml} = xmlStream;
    const iConditionalFormatting = xml.indexOf('conditionalFormatting');
    const iDataValidations = xml.indexOf('dataValidations');
    expect(iConditionalFormatting).not.toBe(-1);
    expect(iDataValidations).not.toBe(-1);
    expect(iConditionalFormatting).toBeLessThan(iDataValidations);
  });
});
