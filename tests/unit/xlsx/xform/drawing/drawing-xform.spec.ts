

import fs from 'fs';
import testXformHelper from '../test-xform-helper.ts';
import _initialModel from './data/drawing.1.0.ts';
import _preparedModel from './data/drawing.1.1.ts';
import _parsedModel from './data/drawing.1.3.ts';
import _reconciledModel from './data/drawing.1.4.ts';

import DrawingXform from '../../../../../src/xlsx/xform/drawing/drawing-xform.ts';
const options = {
  rels: {
    rId1: {Target: '../media/image1.jpg'},
    rId2: {Target: '../media/image2.jpg'},
  },
  mediaIndex: {image1: 0, image2: 1},
  media: [{}, {}],
};

const expectations = [
  {
    title: 'Drawing 1',
    create() {
      return new DrawingXform();
    },
    initialModel: _initialModel,
    preparedModel: _preparedModel,
    xml: fs.readFileSync(`${__dirname}/data/drawing.1.2.xml`).toString(),
    parsedModel: _parsedModel,
    reconciledModel: _reconciledModel,
    tests: ['prepare', 'render', 'renderIn', 'parse', 'reconcile'],
    options,
  },
];

describe('DrawingXform', () => {
  testXformHelper(expectations);
});
