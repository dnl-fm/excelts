
import testXformHelper from '../../test-xform-helper.ts';

import FormulaXform from '../../../../../../src/xlsx/xform/sheet/cf/formula-xform.ts';
const expectations = [
  {
    title: 'formula',
    create() {
      return new FormulaXform();
    },
    preparedModel: 'ROW()',
    xml: '<formula>ROW()</formula>',
    parsedModel: 'ROW()',
    tests: ['render', 'parse'],
  },
];

describe('FormulaXform', () => {
  testXformHelper(expectations);
});
