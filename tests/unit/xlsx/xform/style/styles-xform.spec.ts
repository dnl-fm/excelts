

import fs from 'fs';
import testXformHelper from '../test-xform-helper.ts';
import { expectXmlEqual } from '../test-xform-helper.ts';
import _preparedModel from './data/styles.1.1.json';

import StylesXform from '../../../../../src/xlsx/xform/style/styles-xform.ts';
import XmlStream from '../../../../../src/utils/xml-stream.ts';
const expectations = [
  {
    title: 'Styles with fonts',
    create() {
      return new StylesXform();
    },
    preparedModel: _preparedModel,
    xml: fs.readFileSync(`${__dirname}/data/styles.1.2.xml`).toString(),
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('StylesXform', () => {
  testXformHelper(expectations);

  describe('As StyleManager', () => {
    it('Renders empty model', () => {
      const stylesXform = new StylesXform(true);
      const expectedXml = fs
        .readFileSync(`${__dirname}/data/styles.2.2.xml`)
        .toString();

      const xmlStream = new XmlStream();
      stylesXform.render(xmlStream);

      expectXmlEqual(xmlStream.xml, expectedXml);
    });
  });
});
