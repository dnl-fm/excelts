import BaseXform from '../base-xform.ts';
import CacheField from './cache-field.ts';
import XmlStream from '../../../utils/xml-stream.ts';
import type {XmlStreamWriter, SaxNode} from '../xform-types.ts';

type SourceSheet = {
  name: string;
  dimensions: {
    shortRange: string;
  };
};

type CacheFieldData = {
  name: string;
  sharedItems: string[] | null;
};

type PivotCacheDefinitionModel = {
  sourceSheet: SourceSheet;
  cacheFields: CacheFieldData[];
};

class PivotCacheDefinitionXform extends BaseXform {
  static PIVOT_CACHE_DEFINITION_ATTRIBUTES: Record<string, string> = {
    xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'mc:Ignorable': 'xr',
    'xmlns:xr': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
  };

  constructor() {
    super();

    this.map = {};
  }

  prepare(_model: unknown) {
    // TK
  }

  get tag() {
    // http://www.datypic.com/sc/ooxml/e-ssml_pivotCacheDefinition.html
    return 'pivotCacheDefinition';
  }

  render(xmlStream: XmlStreamWriter, model: PivotCacheDefinitionModel) {
    const {sourceSheet, cacheFields} = model;

    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode(this.tag, {
      ...PivotCacheDefinitionXform.PIVOT_CACHE_DEFINITION_ATTRIBUTES,
      'r:id': 'rId1',
      refreshOnLoad: '1', // important for our implementation to work
      refreshedBy: 'Author',
      refreshedDate: '45125.026046874998',
      createdVersion: '8',
      refreshedVersion: '8',
      minRefreshableVersion: '3',
      recordCount: cacheFields.length + 1,
    });

    xmlStream.openNode('cacheSource', {type: 'worksheet'});
    xmlStream.leafNode('worksheetSource', {
      ref: sourceSheet.dimensions.shortRange,
      sheet: sourceSheet.name,
    });
    xmlStream.closeNode();

    xmlStream.openNode('cacheFields', {count: cacheFields.length});
    // Note: keeping this pretty-printed for now to ease debugging.
    xmlStream.writeXml(cacheFields.map(cacheField => new CacheField(cacheField).render()).join('\n    '));
    xmlStream.closeNode();

    xmlStream.closeNode();
  }

  parseOpen(_node: SaxNode) {
    // TK
    return false;
  }

  parseText(_text: string) {
    // TK
  }

  parseClose(_name: string): boolean {
    // TK
    return false;
  }

  reconcile(_model: unknown, _options: unknown) {
    // TK
  }
}

export default PivotCacheDefinitionXform;
