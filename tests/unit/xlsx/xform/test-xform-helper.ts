
import CompyXform from './compy-xform.ts';
import { PassThrough } from 'stream';
import _ from '../../../utils/under-dash.ts';

import parseSax from '../../../../src/utils/parse-sax.ts';
import XmlStream from '../../../../src/utils/xml-stream.ts';
import BooleanXform from '../../../../src/xlsx/xform/simple/boolean-xform.ts';
const { cloneDeep, each } = _;

function getExpectation(expectation: any, name: string) {
  if (!expectation.hasOwnProperty(name)) {
    throw new Error(`Expectation missing required field: ${name}`);
  }
  return cloneDeep(expectation[name]);
}

/**
 * Normalize XML for comparison (whitespace-insensitive, attribute-order-insensitive)
 */
function normalizeXml(xml: string): string {
  // First pass: basic whitespace normalization
  let result = xml
    .replace(/>\s+</g, '><')
    .replace(/\s+/g, ' ')
    .replace(/<([\w:]+)([^>]*?)\s*\/>/g, '<$1$2></$1>') // Convert self-closing to explicit close
    .replace(/\s+>/g, '>') // Remove whitespace before >
    .trim();

  // Second pass: sort attributes within each tag for order-insensitive comparison
  // Handle namespaced tags like x14:dataBar with [\w:]+
  result = result.replace(/<([\w:]+)((?:\s+[^>]+)?)>/g, (match, tagName, attrs) => {
    if (!attrs || !attrs.trim()) {
      return `<${tagName}>`;
    }
    // Parse attributes
    const attrRegex = /(\w+(?::\w+)?)\s*=\s*"([^"]*)"/g;
    const attrList: [string, string][] = [];
    let m;
    while ((m = attrRegex.exec(attrs)) !== null) {
      attrList.push([m[1], m[2]]);
    }
    // Sort by attribute name
    attrList.sort((a, b) => a[0].localeCompare(b[0]));
    const sortedAttrs = attrList.map(([k, v]) => `${k}="${v}"`).join(' ');
    return `<${tagName}${sortedAttrs ? ' ' + sortedAttrs : ''}>`;
  });

  return result;
}

/**
 * Compare XML strings
 */
function expectXmlEqual(actual: string, expected: string): void {
  expect(normalizeXml(actual)).toBe(normalizeXml(expected));
}

// ===============================================================================================================
// provides boilerplate examples for the four transform steps: prepare, render, parse and reconcile
//  prepare: model => preparedModel
//  render:  preparedModel => xml
//  parse:  xml => parsedModel
//  reconcile: parsedModel => reconciledModel

const its: Record<string, (expectation: any) => void> = {
  prepare(expectation) {
    it('Prepare Model', () =>
      new Promise<void>(resolve => {
        const model = getExpectation(expectation, 'initialModel');
        const result = getExpectation(expectation, 'preparedModel');

        const xform = expectation.create();
        xform.prepare(model, expectation.options);
        expect(cloneDeep(model, false)).toEqual(result);
        resolve();
      }));
  },

  render(expectation) {
    it('Render to XML', () =>
      new Promise<void>(resolve => {
        const model = getExpectation(expectation, 'preparedModel');
        const result = getExpectation(expectation, 'xml');

        const xform = expectation.create();
        const xmlStream = new XmlStream();
        xform.render(xmlStream, model, 0);

        expectXmlEqual(xmlStream.xml, result);
        resolve();
      }));
  },

  'prepare-render': function(expectation) {
    // when implementation details get in the way of testing the prepared result
    it('Prepare and Render to XML', () =>
      new Promise<void>(resolve => {
        const model = getExpectation(expectation, 'initialModel');
        const result = getExpectation(expectation, 'xml');

        const xform = expectation.create();
        const xmlStream = new XmlStream();

        xform.prepare(model, expectation.options);
        xform.render(xmlStream, model);

        expectXmlEqual(xmlStream.xml, result);
        resolve();
      }));
  },

  renderIn(expectation) {
    it('Render in Composite to XML ', () =>
      new Promise<void>(resolve => {
        const model = {
          pre: true,
          child: getExpectation(expectation, 'preparedModel'),
          post: true,
        };
        const result = `<compy><pre/>${getExpectation(
          expectation,
          'xml'
        )}<post/></compy>`;

        const xform = new CompyXform({
          tag: 'compy',
          children: [
            {
              name: 'pre',
              xform: new BooleanXform({tag: 'pre', attr: 'val'}),
            },
            {name: 'child', xform: expectation.create()},
            {
              name: 'post',
              xform: new BooleanXform({tag: 'post', attr: 'val'}),
            },
          ],
        });

        const xmlStream = new XmlStream();
        xform.render(xmlStream, model);

        expectXmlEqual(xmlStream.xml, result);
        resolve();
      }));
  },

  parseIn(expectation) {
    it('Parse within composite', () =>
      new Promise<void>((resolve, reject) => {
        const xml = `<compy><pre/>${getExpectation(
          expectation,
          'xml'
        )}<post/></compy>`;
        const childXform = expectation.create();
        const result: Record<string, any> = {pre: true};
        result[childXform.tag] = getExpectation(expectation, 'parsedModel');
        result.post = true;
        const xform = new CompyXform({
          tag: 'compy',
          children: [
            {
              name: 'pre',
              xform: new BooleanXform({tag: 'pre', attr: 'val'}),
            },
            {name: childXform.tag, xform: childXform},
            {
              name: 'post',
              xform: new BooleanXform({tag: 'post', attr: 'val'}),
            },
          ],
        });
        const stream = new PassThrough();
        stream.write(xml);
        stream.end();
        xform
          .parse(parseSax(stream))
          .then((model: any) => {
            const clone = cloneDeep(model, false);
            expect(clone).toEqual(result);
            resolve();
          })
          .catch(reject);
      }));
  },

  parse(expectation) {
    it('Parse to Model', () =>
      new Promise<void>((resolve, reject) => {
        const xml = getExpectation(expectation, 'xml');
        const result = getExpectation(expectation, 'parsedModel');

        const xform = expectation.create();

        const stream = new PassThrough();
        stream.write(xml);
        stream.end();
        xform
          .parse(parseSax(stream))
          .then((model: any) => {
            const clone = cloneDeep(model, false);
            expect(clone).toEqual(result);
            resolve();
          })
          .catch(reject);
      }));
  },

  reconcile(expectation) {
    it('Reconcile Model', () =>
      new Promise<void>(resolve => {
        const model = getExpectation(expectation, 'parsedModel');
        const result = getExpectation(expectation, 'reconciledModel');

        const xform = expectation.create();
        xform.reconcile(model, expectation.options);

        const clone = cloneDeep(model, false);
        expect(clone).toEqual(result);
        resolve();
      }));
  },
};

function testXformHelper(expectations: any[]) {
  each(expectations, (expectation: any) => {
    const tests = getExpectation(expectation, 'tests');
    describe(expectation.title, () => {
      each(tests, (test: string) => {
        its[test](expectation);
      });
    });
  });
}

export { expectXmlEqual, normalizeXml };
export default testXformHelper;
