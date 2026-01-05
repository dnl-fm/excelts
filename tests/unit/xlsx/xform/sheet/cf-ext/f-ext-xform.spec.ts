
import testXformHelper from '../../test-xform-helper.ts';

import FExtXform from '../../../../../../src/xlsx/xform/sheet/cf-ext/f-ext-xform.ts';
const expectations = [
  {
    title: 'formula',
    create() {
      return new FExtXform();
    },
    preparedModel: '7',
    xml: '<xm:f>7</xm:f>',
    parsedModel: '7',
    tests: ['render', 'parse'],
  },
];

describe('FExtXform', () => {
  testXformHelper(expectations);
});
