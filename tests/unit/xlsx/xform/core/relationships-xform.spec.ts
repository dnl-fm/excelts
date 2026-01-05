

import fs from 'fs';
import testXformHelper from '../test-xform-helper.ts';
import _preparedModel from './data/worksheet.rels.1.json';

const RelationshipsXform = verquire('xlsx/xform/core/relationships-xform');

const expectations = [
  {
    title: 'worksheet.rels',
    create() {
      return new RelationshipsXform();
    },
    preparedModel: _preparedModel,
    xml: fs
      .readFileSync(`${__dirname}/data/worksheet.rels.xml`)
      .toString()
      .replace(/\r\n/g, '\n'),
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('RelationshipsXform', () => {
  testXformHelper(expectations);
});
