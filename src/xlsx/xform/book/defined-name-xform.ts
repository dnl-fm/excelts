
import BaseXform from '../base-xform.ts';
import colCache from '../../../utils/col-cache.ts';

interface DefinedNameModel {
  name: string;
  ranges: string[];
  localSheetId?: number;
}

class DefinedNamesXform extends BaseXform {
  _parsedName?: string;
  _parsedLocalSheetId?: string;
  _parsedText: string[] = [];
  declare model: DefinedNameModel | undefined;

  render(xmlStream: any, model: DefinedNameModel) {
    // <definedNames>
    //   <definedName name="name">name.ranges.join(',')</definedName>
    //   <definedName name="_xlnm.Print_Area" localSheetId="0">name.ranges.join(',')</definedName>
    // </definedNames>
    xmlStream.openNode('definedName', {
      name: model.name,
      localSheetId: model.localSheetId,
    });
    xmlStream.writeText(model.ranges.join(','));
    xmlStream.closeNode();
  }

  parseOpen(node: { name: string; attributes: Record<string, string> }) {
    switch (node.name) {
      case 'definedName':
        this._parsedName = node.attributes.name;
        this._parsedLocalSheetId = node.attributes.localSheetId;
        this._parsedText = [];
        return true;
      default:
        return false;
    }
  }

  parseText(text: string) {
    this._parsedText.push(text);
  }

  parseClose(_name: string) {
    this.model = {
      name: this._parsedName!,
      ranges: extractRanges(this._parsedText.join('')),
    };
    if (this._parsedLocalSheetId !== undefined) {
      this.model.localSheetId = parseInt(this._parsedLocalSheetId, 10);
    }
    return false;
  }
}

function isValidRange(range: string) {
  try {
    colCache.decodeEx(range);
    return true;
  } catch {
    return false;
  }
}

function extractRanges(parsedText: string): string[] {
  const ranges = [];
  let quotesOpened = false;
  let last = '';
  parsedText.split(',').forEach(item => {
    if (!item) {
      return;
    }
    const quotes = (item.match(/'/g) || []).length;

    if (!quotes) {
      if (quotesOpened) {
        last += `${item},`;
      } else if (isValidRange(item)) {
        ranges.push(item);
      }
      return;
    }
    const quotesEven = quotes % 2 === 0;

    if (!quotesOpened && quotesEven && isValidRange(item)) {
      ranges.push(item);
    } else if (quotesOpened && !quotesEven) {
      quotesOpened = false;
      if (isValidRange(last + item)) {
        ranges.push(last + item);
      }
      last = '';
    } else {
      quotesOpened = true;
      last += `${item},`;
    }
  });
  return ranges;
}

export default DefinedNamesXform;