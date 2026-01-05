

import fs from 'fs';
import testXformHelper from '../test-xform-helper.ts';
import _preparedModel from './data/table.1.1.json';
import _parsedModel from './data/table.1.3.json';

import TableXform from '../../../../../src/xlsx/xform/table/table-xform.ts';
const expectations = [
  {
    title: 'showing filter',
    create() {
      return new TableXform();
    },
    initialModel: null,
    preparedModel: _preparedModel,
    xml: fs.readFileSync(`${__dirname}/data/table.1.2.xml`).toString(),
    parsedModel: _parsedModel,
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('TableXform', () => {
  testXformHelper(expectations);
});
