
import testXformHelper from '../../test-xform-helper.ts';

import SqrefExtXform from '../../../../../../src/xlsx/xform/sheet/cf-ext/sqref-ext-xform.ts';
const expectations = [
  {
    title: 'range',
    create() {
      return new SqrefExtXform();
    },
    preparedModel: 'A1:C3',
    xml: '<xm:sqref>A1:C3</xm:sqref>',
    parsedModel: 'A1:C3',
    tests: ['render', 'parse'],
  },
];

describe('SqrefExtXform', () => {
  testXformHelper(expectations);
});
