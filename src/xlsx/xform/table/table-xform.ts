import XmlStream from '../../../utils/xml-stream.ts';
import BaseXform from '../base-xform.ts';
import ListXform from '../list-xform.ts';
import AutoFilterXform from './auto-filter-xform.ts';
import TableColumnXform from './table-column-xform.ts';
import TableStyleInfoXform from './table-style-info-xform.ts';
import type {XmlStreamWriter, SaxNode} from '../xform-types.ts';

type TableColumnModel = {
  name?: string;
  dxfId?: number;
  filterButton?: boolean;
  style?: unknown;
};

type TableStyleModel = {
  name?: string;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
};

type AutoFilterModel = {
  autoFilterRef?: string;
  columns?: Array<{filterButton?: boolean}>;
};

type TableModel = {
  id?: number;
  name?: string;
  displayName?: string;
  tableRef?: string;
  totalsRow?: boolean;
  headerRow?: boolean;
  columns?: TableColumnModel[];
  style?: TableStyleModel;
  autoFilterRef?: string;
};

type TableXformPrepareOptions = {
  tableId?: number;
};

type TableXformReconcileOptions = {
  styles?: {
    getDxfStyle: (id: number) => unknown;
  };
};

class TableXform extends BaseXform {
  static TABLE_ATTRIBUTES: Record<string, string> = {
    xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'mc:Ignorable': 'xr xr3',
    'xmlns:xr': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
    'xmlns:xr3': 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision3',
    // 'xr:uid': '{00000000-000C-0000-FFFF-FFFF00000000}',
  };

  declare model: TableModel | undefined;
  declare map: {
    autoFilter: AutoFilterXform;
    tableColumns: ListXform;
    tableStyleInfo: TableStyleInfoXform;
  };

  constructor() {
    super();

    this.map = {
      autoFilter: new AutoFilterXform(),
      tableColumns: new ListXform({
        tag: 'tableColumns',
        count: true,
        empty: true,
        childXform: new TableColumnXform(),
      }),
      tableStyleInfo: new TableStyleInfoXform(),
    };
  }

  prepare(model: TableModel, options: TableXformPrepareOptions) {
    this.map.autoFilter.prepare(model);
    this.map.tableColumns.prepare(model.columns, options);
  }

  get tag() {
    return 'table';
  }

  render(xmlStream: XmlStreamWriter, model: TableModel) {
    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode(this.tag, {
      ...TableXform.TABLE_ATTRIBUTES,
      id: model.id,
      name: model.name,
      displayName: model.displayName || model.name,
      ref: model.tableRef,
      totalsRowCount: model.totalsRow ? '1' : undefined,
      totalsRowShown: model.totalsRow ? undefined : '1',
      headerRowCount: model.headerRow ? '1' : '0',
    });

    this.map.autoFilter.render(xmlStream, model);
    this.map.tableColumns.render(xmlStream, model.columns);
    this.map.tableStyleInfo.render(xmlStream, model.style);

    xmlStream.closeNode();
  }

  parseOpen(node: SaxNode) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    const {name, attributes} = node;
    switch (name) {
      case this.tag:
        this.reset();
        this.model = {
          name: attributes.name as string,
          displayName: (attributes.displayName || attributes.name) as string,
          tableRef: attributes.ref as string,
          totalsRow: attributes.totalsRowCount === '1',
          headerRow: attributes.headerRowCount === '1',
        };
        break;
      default:
        this.parser = this.map[node.name as keyof typeof this.map];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  }

  parseText(text: string) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        this.model!.columns = this.map.tableColumns.model as TableColumnModel[];
        if (this.map.autoFilter.model) {
          const autoFilterModel = this.map.autoFilter.model as AutoFilterModel;
          this.model!.autoFilterRef = autoFilterModel.autoFilterRef;
          autoFilterModel.columns?.forEach((column, index) => {
            this.model!.columns![index].filterButton = column.filterButton;
          });
        }
        this.model!.style = this.map.tableStyleInfo.model as TableStyleModel;
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  }

  reconcile(model: TableModel, options: TableXformReconcileOptions) {
    // fetch the dfxs from styles
    model.columns?.forEach(column => {
      if (column.dxfId !== undefined && options.styles) {
        column.style = options.styles.getDxfStyle(column.dxfId);
      }
    });
  }
}

export default TableXform;
