

import fs from 'fs';
import testXformHelper from '../test-xform-helper.ts';
import _preparedModel from './data/sharedStrings.json';

import SharedStringsXform from '../../../../../src/xlsx/xform/strings/shared-strings-xform.ts';
const expectations = [
  {
    title: 'Shared Strings',
    create() {
      return new SharedStringsXform();
    },
    preparedModel: _preparedModel,
    xml: fs.readFileSync(`${__dirname}/data/sharedStrings.xml`).toString(),
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('SharedStringsXform', () => {
  testXformHelper(expectations);
});
