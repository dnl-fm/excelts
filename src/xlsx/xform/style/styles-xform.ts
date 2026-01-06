/* eslint-disable max-classes-per-file */

import Enums from '../../../doc/enums.ts';
import XmlStream from '../../../utils/xml-stream.ts';
import BaseXform from '../base-xform.ts';
import StaticXform from '../static-xform.ts';
import ListXform from '../list-xform.ts';
import FontXform from './font-xform.ts';
import FillXform from './fill-xform.ts';
import BorderXform from './border-xform.ts';
import NumFmtXform from './numfmt-xform.ts';
import StyleXform from './style-xform.ts';
import DxfXform from './dxf-xform.ts';

export interface StyleModel {
  numFmtId?: number;
  fontId?: number;
  fillId?: number;
  borderId?: number;
  xfId?: number;
  [key: string]: unknown;
}

export interface NumFmtModel {
  id?: number;
  formatCode?: string;
}

export interface StylesModel {
  styles: unknown[];
  numFmts: (NumFmtModel | string)[];
  fonts: unknown[];
  borders: unknown[];
  fills: unknown[];
  dxfs: unknown[];
  [key: string]: unknown;
}

const NUMFMT_BASE = 164;

class StylesXform extends BaseXform {
  declare model: StylesModel | undefined;
  index?: Record<string, unknown>;
  weakMap?: WeakMap<object, unknown>;
  _dateStyleId?: number;

  static STYLESHEET_ATTRIBUTES: Record<string, string>;
  static STATIC_XFORMS: Record<string, BaseXform>;
  static Mock: typeof StylesXformMock;

  constructor(initialise?: boolean) {
    super();

    this.map = {
      numFmts: new ListXform({tag: 'numFmts', count: true, childXform: new NumFmtXform()}),
      fonts: new ListXform({
        tag: 'fonts',
        count: true,
        childXform: new FontXform(),
        $: {'x14ac:knownFonts': 1},
      }),
      fills: new ListXform({tag: 'fills', count: true, childXform: new FillXform()}),
      borders: new ListXform({tag: 'borders', count: true, childXform: new BorderXform()}),
      cellStyleXfs: new ListXform({tag: 'cellStyleXfs', count: true, childXform: new StyleXform()}),
      cellXfs: new ListXform({
        tag: 'cellXfs',
        count: true,
        childXform: new StyleXform({xfId: true}),
      }),
      dxfs: new ListXform({tag: 'dxfs', always: true, count: true, childXform: new DxfXform()}),

      // for style manager
      numFmt: new NumFmtXform(),
      font: new FontXform(),
      fill: new FillXform(),
      border: new BorderXform(),
      style: new StyleXform({xfId: true}),

      cellStyles: StylesXform.STATIC_XFORMS.cellStyles,
      tableStyles: StylesXform.STATIC_XFORMS.tableStyles,
      extLst: StylesXform.STATIC_XFORMS.extLst,
    };

    if (initialise) {
      // StylesXform also acts as style manager and is used to build up styles-model during worksheet processing
      this.init();
    }
  }

  initIndex() {
    this.index = {
      style: {},
      numFmt: {},
      numFmtNextId: 164, // start custom format ids here
      font: {},
      border: {},
      fill: {},
    };
  }

  init() {
    // Prepare for Style Manager role
    this.model = {
      styles: [],
      numFmts: [],
      fonts: [],
      borders: [],
      fills: [],
      dxfs: [],
    };

    this.initIndex();

    // default (zero) border
    this._addBorder({});

    // add default (all zero) style
    this._addStyle({numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0});

    // add default fills
    this._addFill({type: 'pattern', pattern: 'none'});
    this._addFill({type: 'pattern', pattern: 'gray125'});

    this.weakMap = new WeakMap();
  }

  render(xmlStream: unknown, model?: unknown): void {
    model = model || this.model;
    //
    //   <fonts count="2" x14ac:knownFonts="1">
    (xmlStream as any).openXml(XmlStream.StdDocAttributes);

    (xmlStream as any).openNode('styleSheet', StylesXform.STYLESHEET_ATTRIBUTES);

    if (this.index) {
      // model has been built by style manager role (contains xml)
      const m = model as StylesModel;
      if (m.numFmts && m.numFmts.length) {
        (xmlStream as any).openNode('numFmts', {count: m.numFmts.length});
        m.numFmts.forEach(numFmtXml => {
          (xmlStream as any).writeXml(numFmtXml);
        });
        (xmlStream as any).closeNode();
      }

      if (!m.fonts.length) {
        // default (zero) font
        this._addFont({size: 11, color: {theme: 1}, name: 'Calibri', family: 2, scheme: 'minor'});
      }
      (xmlStream as any).openNode('fonts', {count: m.fonts.length, 'x14ac:knownFonts': 1});
      m.fonts.forEach(fontXml => {
        (xmlStream as any).writeXml(fontXml);
      });
      (xmlStream as any).closeNode();

      (xmlStream as any).openNode('fills', {count: m.fills.length});
      m.fills.forEach(fillXml => {
        (xmlStream as any).writeXml(fillXml);
      });
      (xmlStream as any).closeNode();

      (xmlStream as any).openNode('borders', {count: m.borders.length});
      m.borders.forEach(borderXml => {
        (xmlStream as any).writeXml(borderXml);
      });
      (xmlStream as any).closeNode();

      (this.map!.cellStyleXfs as BaseXform).render(xmlStream, [{numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0}]);

      (xmlStream as any).openNode('cellXfs', {count: m.styles.length});
      m.styles.forEach(styleXml => {
        (xmlStream as any).writeXml(styleXml);
      });
      (xmlStream as any).closeNode();
    } else {
      // model is plain JSON and needs to be xformed
      const m = model as StylesModel;
      (this.map!.numFmts as BaseXform).render(xmlStream, m.numFmts);
      (this.map!.fonts as BaseXform).render(xmlStream, m.fonts);
      (this.map!.fills as BaseXform).render(xmlStream, m.fills);
      (this.map!.borders as BaseXform).render(xmlStream, m.borders);
      (this.map!.cellStyleXfs as BaseXform).render(xmlStream, [{numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0}]);
      (this.map!.cellXfs as BaseXform).render(xmlStream, m.styles);
    }

    StylesXform.STATIC_XFORMS.cellStyles.render(xmlStream);

    (this.map!.dxfs as BaseXform).render(xmlStream, (model as StylesModel).dxfs);

    StylesXform.STATIC_XFORMS.tableStyles.render(xmlStream);
    StylesXform.STATIC_XFORMS.extLst.render(xmlStream);

    (xmlStream as any).closeNode();
  }

  parseOpen(node: unknown): boolean {
    if ((this as any).parser) {
      ((this as any).parser as BaseXform).parseOpen(node);
      return true;
    }
    const nodeName = (node as any).name;
    switch (nodeName) {
      case 'styleSheet':
        this.initIndex();
        return true;
      default:
        this.parser = this.map![nodeName] as BaseXform | undefined;
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        return true;
    }
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case 'styleSheet': {
        this.model = {} as StylesModel;
        const add = (propName: string, xform: BaseXform) => {
          const xformModel = (xform as any)?.model;
          if (xformModel?.length) {
            (this.model as any)[propName] = xformModel;
          }
        };
        add('numFmts', this.map!.numFmts as BaseXform);
        add('fonts', this.map!.fonts as BaseXform);
        add('fills', this.map!.fills as BaseXform);
        add('borders', this.map!.borders as BaseXform);
        add('styles', this.map!.cellXfs as BaseXform);
        add('dxfs', this.map!.dxfs as BaseXform);

        // index numFmts
        this.index = {
          model: [] as unknown[],
          numFmt: {} as Record<number | string, unknown>,
        };
        if (this.model.numFmts) {
          const numFmtIndex = this.index.numFmt;
          (this.model?.numFmts || []).forEach((numFmt: NumFmtModel) => {
            numFmtIndex[numFmt.id!] = numFmt.formatCode;
          });
        }

        return false;
      }
      default:
        // not quite sure how we get here!
        return true;
    }
  }

  // add a cell's style model to the collection
  // each style property is processed and cross-referenced, etc.
  // the styleId is returned. Note: cellType is used when numFmt not defined
  addStyleModel(model: unknown, cellType?: unknown): number {
    if (!model) {
      return 0;
    }

    // if we have no default font, add it here now
    if (!this.model!.fonts.length) {
      // default (zero) font
      this._addFont({size: 11, color: {theme: 1}, name: 'Calibri', family: 2, scheme: 'minor'});
    }

    // if we have seen this style object before, assume it has the same styleId
    if (this.weakMap && this.weakMap.has(model as object)) {
      return this.weakMap.get(model as object) as number;
    }

    const style: StyleModel = {};
    cellType = cellType || Enums.ValueType.Number;

    const m = model as Record<string, unknown>;
    if (m.numFmt) {
      style.numFmtId = this._addNumFmtStr(m.numFmt as string);
    } else {
      switch (cellType) {
        case Enums.ValueType.Number:
          style.numFmtId = this._addNumFmtStr('General');
          break;
        case Enums.ValueType.Date:
          style.numFmtId = this._addNumFmtStr('mm-dd-yy');
          break;
        default:
          break;
      }
    }

    if (m.font) {
      style.fontId = this._addFont(m.font);
    }

    if (m.border) {
      style.borderId = this._addBorder(m.border);
    }

    if (m.fill) {
      style.fillId = this._addFill(m.fill);
    }

    if (m.alignment) {
      style.alignment = m.alignment;
    }

    if (m.protection) {
      style.protection = m.protection;
    }

    const styleId = this._addStyle(style);
    if (this.weakMap) {
      this.weakMap.set(model as object, styleId);
    }
    return styleId;
  }

  // given a styleId (i.e. s="n"), get the cell's style model
  // objects are shared where possible.
  getStyleModel(id: number): unknown {
    const style = (this.model?.styles?.[id]) as StyleModel | undefined;
    if (!style) return null;

    // have we built this model before?
    let model = this.index!.model![id] as Record<string, unknown> | undefined;
    if (model) return model;

    // build a new model
    model = this.index!.model![id] = {} as Record<string, unknown>;

    // -------------------------------------------------------
    // number format
    if (style.numFmtId) {
      const numFmt = this.index!.numFmt![style.numFmtId] || NumFmtXform.getDefaultFmtCode(style.numFmtId);
      if (numFmt) {
        model.numFmt = numFmt;
      }
    }

    function addStyle(name: string, group: unknown, styleId: number | undefined) {
      if (styleId !== undefined) {
        const part = (group as any)?.[styleId];
        if (part) {
          model![name] = part;
        }
      }
    }

    addStyle('font', this.model!.fonts, style.fontId);
    addStyle('border', this.model!.borders, style.borderId);
    addStyle('fill', this.model!.fills, style.fillId);

    // -------------------------------------------------------
    // alignment
    if (style.alignment) {
      model.alignment = style.alignment;
    }

    // -------------------------------------------------------
    // protection
    if (style.protection) {
      model.protection = style.protection;
    }

    return model;
  }

  addDxfStyle(style: unknown): number {
    const s = style as Record<string, unknown>;
    if (s.numFmt) {
      // register numFmtId to use it during dxf-xform rendering
      s.numFmtId = this._addNumFmtStr(s.numFmt as string);
    }

    this.model!.dxfs.push(style);
    return this.model!.dxfs.length - 1;
  }

  getDxfStyle(id: number): unknown {
    return this.model?.dxfs?.[id];
  }

  // =========================================================================
  // Private Interface
  _addStyle(style: unknown): number {
    const xml = (this.map!.style as BaseXform).toXml(style);
    let index = (this.index!.style as Record<string, number>)[xml];
    if (index === undefined) {
      index = (this.index!.style as Record<string, number>)[xml] = this.model!.styles.length;
      this.model!.styles.push(xml);
    }
    return index;
  }

  // =========================================================================
  // Number Formats
  _addNumFmtStr(formatCode: string): number {
    // check if default format
    let index = NumFmtXform.getDefaultFmtId(formatCode);
    if (index !== undefined) return index;

    // check if already in
    index = (this.index!.numFmt as Record<string, number>)[formatCode];
    if (index !== undefined) return index;

    index = (this.index!.numFmt as Record<string, number>)[formatCode] = NUMFMT_BASE + this.model!.numFmts.length;
    const xml = (this.map!.numFmt as BaseXform).toXml({id: index, formatCode});
    this.model!.numFmts.push(xml);
    return index;
  }

  // =========================================================================
  // Fonts
  _addFont(font: unknown): number {
    const xml = (this.map!.font as BaseXform).toXml(font);
    let index = (this.index!.font as Record<string, number>)[xml];
    if (index === undefined) {
      index = (this.index!.font as Record<string, number>)[xml] = this.model!.fonts.length;
      this.model!.fonts.push(xml);
    }
    return index;
  }

  // =========================================================================
  // Borders
  _addBorder(border: unknown): number {
    const xml = (this.map!.border as BaseXform).toXml(border);
    let index = (this.index!.border as Record<string, number>)[xml];
    if (index === undefined) {
      index = (this.index!.border as Record<string, number>)[xml] = this.model!.borders.length;
      this.model!.borders.push(xml);
    }
    return index;
  }

  // =========================================================================
  // Fills
  _addFill(fill: unknown): number {
    const xml = (this.map!.fill as BaseXform).toXml(fill);
    let index = (this.index!.fill as Record<string, number>)[xml];
    if (index === undefined) {
      index = (this.index!.fill as Record<string, number>)[xml] = this.model!.fills.length;
      this.model!.fills.push(xml);
    }
    return index;
  }

  // =========================================================================
}

StylesXform.STYLESHEET_ATTRIBUTES = {
  xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  'mc:Ignorable': 'x14ac x16r2',
  'xmlns:x14ac': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
  'xmlns:x16r2': 'http://schemas.microsoft.com/office/spreadsheetml/2015/02/main',
};
StylesXform.STATIC_XFORMS = {
  cellStyles: new StaticXform({
    tag: 'cellStyles',
    $: {count: 1},
    c: [{tag: 'cellStyle', $: {name: 'Normal', xfId: 0, builtinId: 0}}],
  }),
  dxfs: new StaticXform({tag: 'dxfs', $: {count: 0}}),
  tableStyles: new StaticXform({
    tag: 'tableStyles',
    $: {count: 0, defaultTableStyle: 'TableStyleMedium2', defaultPivotStyle: 'PivotStyleLight16'},
  }),
  extLst: new StaticXform({
    tag: 'extLst',
    c: [
      {
        tag: 'ext',
        $: {
          uri: '{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}',
          'xmlns:x14': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
        },
        c: [{tag: 'x14:slicerStyles', $: {defaultSlicerStyle: 'SlicerStyleLight1'}}],
      },
      {
        tag: 'ext',
        $: {
          uri: '{9260A510-F301-46a8-8635-F512D64BE5F5}',
          'xmlns:x15': 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main',
        },
        c: [{tag: 'x15:timelineStyles', $: {defaultTimelineStyle: 'TimeSlicerStyleLight1'}}],
      },
    ],
  }),
};

// the stylemanager mock acts like StyleManager except that it always returns 0 or {}
class StylesXformMock extends StylesXform {
  constructor(_initialise?: boolean) {
    super();

    this.model = {
      styles: [{numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0}],
      numFmts: [],
      fonts: [{size: 11, color: {theme: 1}, name: 'Calibri', family: 2, scheme: 'minor'}],
      borders: [{}],
      fills: [
        {type: 'pattern', pattern: 'none'},
        {type: 'pattern', pattern: 'gray125'},
      ],
      dxfs: [],
    };
  }

  // =========================================================================
  // Style Manager Interface

  // override normal behaviour - consume and dispose
  parseStream(stream: unknown): Promise<void> {
    (stream as any).autodrain();
    return Promise.resolve();
  }

  // add a cell's style model to the collection
  // each style property is processed and cross-referenced, etc.
  // the styleId is returned. Note: cellType is used when numFmt not defined
  addStyleModel(_model: unknown, cellType?: unknown): number {
    switch (cellType) {
      case Enums.ValueType.Date:
        return this.dateStyleId;
      default:
        return 0;
    }
  }

  get dateStyleId(): number {
    if (!this._dateStyleId) {
      const dateStyle = {
        numFmtId: NumFmtXform.getDefaultFmtId('mm-dd-yy'),
      };
      this._dateStyleId = this.model!.styles.length;
      this.model!.styles.push(dateStyle);
    }
    return this._dateStyleId;
  }

  // given a styleId (i.e. s="n"), get the cell's style model
  // objects are shared where possible.
  getStyleModel(/* id */): Record<string, unknown> {
    return {};
  }
}

StylesXform.Mock = StylesXformMock;

export default StylesXform;