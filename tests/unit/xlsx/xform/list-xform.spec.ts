
import testXformHelper from './test-xform-helper.ts';

import ListXform from '../../../../src/xlsx/xform/list-xform.ts';
import IntegerXform from '../../../../src/xlsx/xform/simple/integer-xform.ts';
const expectations = [
  {
    title: 'Tagged',
    create() {
      return new ListXform({
        tag: 'ints',
        childXform: new IntegerXform({tag: 'int', attr: 'val'}),
      });
    },
    preparedModel: [1, 2, 3],
    get parsedModel() {
      return this.preparedModel;
    },
    xml: '<ints><int val="1"/><int val="2"/><int val="3"/></ints>',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Tagged and Counted',
    create() {
      return new ListXform({
        tag: 'ints',
        count: true,
        childXform: new IntegerXform({tag: 'int', attr: 'val'}),
      });
    },
    preparedModel: [1, 2, 3],
    get parsedModel() {
      return this.preparedModel;
    },
    xml: '<ints count="3"><int val="1"/><int val="2"/><int val="3"/></ints>',
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('ListXform', () => {
  testXformHelper(expectations);
});
