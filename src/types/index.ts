/**
 * Core types for ExcelTS XLSX I/O operations.
 *
 * @module
 */

// =============================================================================
// Cell Value Types
// =============================================================================

/**
 * Rich text segment for formatted cell content
 */
export interface RichTextSegment {
  text: string;
  font?: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';
    strike?: boolean;
    color?: { argb?: string; theme?: number; tint?: number };
    name?: string;
    size?: number;
    vertAlign?: 'superscript' | 'subscript';
  };
}

/**
 * Rich text cell value
 */
export interface RichTextCellValue {
  richText: RichTextSegment[];
}

/**
 * Hyperlink cell value
 */
export interface HyperlinkCellValue {
  text: string;
  hyperlink: string;
  tooltip?: string;
}

/**
 * Formula cell value
 */
export interface FormulaCellValue {
  formula: string;
  result?: CellPrimitiveValue | RichTextCellValue | HyperlinkCellValue | ErrorCellValue;
  formulaType?: number;
  sharedFormula?: string;
  ref?: string;
  shareType?: string;
}

/**
 * Error cell value
 */
export interface ErrorCellValue {
  error: string;
}

/**
 * Primitive cell values
 */
export type CellPrimitiveValue = null | number | string | boolean | Date;

/**
 * All possible cell value types that can be read from or written to a cell.
 *
 * @example
 * ```ts
 * // Primitive values
 * cell.value = 42;
 * cell.value = 'Hello';
 * cell.value = new Date();
 * cell.value = true;
 *
 * // Formula
 * cell.value = { formula: 'SUM(A1:A10)', result: 55 };
 *
 * // Hyperlink
 * cell.value = { text: 'Click here', hyperlink: 'https://example.com' };
 *
 * // Rich text
 * cell.value = { richText: [{ text: 'Bold', font: { bold: true } }] };
 *
 * // Error
 * cell.value = { error: '#N/A' };
 * ```
 */
export type CellValueType =
  | CellPrimitiveValue
  | RichTextCellValue
  | HyperlinkCellValue
  | FormulaCellValue
  | ErrorCellValue
  | Record<string, unknown>; // For JSON values and edge cases

// =============================================================================
// Stream Types
// =============================================================================

/**
 * Stream types - async iterable interface for Bun-native streaming
 */
export type ReadableStream = AsyncIterable<Uint8Array | string>;

/**
 * Writable stream interface
 */
export interface XlsxWritableStream {
  write(chunk: Uint8Array | string, callback?: (err?: Error) => void): boolean;
  end(callback?: () => void): void;
  on?(event: string, listener: (...args: unknown[]) => void): void;
}

/** @deprecated Use XlsxWritableStream instead */
export type WritableStream = XlsxWritableStream;

/**
 * XLSX options for read operations
 */
export interface XlsxReadOptions {
  base64?: boolean;
  maxRows?: number;
  maxCols?: number;
  ignoreNodes?: string[];
  [key: string]: unknown;
}

/**
 * XLSX options for write operations
 */
export interface XlsxWriteOptions {
  zip?: unknown;
  useSharedStrings?: boolean;
  useStyles?: boolean;
  [key: string]: unknown;
}

/**
 * File system read options
 */
export type FsReadFileOptions = BufferEncoding | { encoding?: BufferEncoding | null; flag?: string } | null;

/**
 * Streaming read/write model shared across transformers
 */
export interface XlsxModel {
  worksheets: WorksheetItemModel[];
  worksheetHash: Record<string, WorksheetItemModel>;
  worksheetRels: Record<string | number, RelationshipModel[]>;
  themes: Record<string, string>;
  media: MediaModel[];
  mediaIndex: Record<string, number>;
  drawings: Record<string, DrawingModel>;
  drawingRels: Record<string, RelationshipModel[]>;
  comments: Record<string, CommentsModel>;
  tables: Record<string, TableModel>;
  vmlDrawings: Record<string, VmlDrawingModel>;
  pivotTablesRaw: Record<string, string>;
  pivotCacheDefinitionsRaw: Record<string, string>;
  pivotCacheRecordsRaw: Record<string, string>;
  globalRels?: RelationshipModel[];
  sheets?: SheetDefinition[];
  definedNames?: unknown;
  views?: unknown[];
  properties?: Record<string, unknown>;
  calcProperties?: unknown;
  workbookRels?: RelationshipModel[];
  sharedStrings?: SharedStringsModel;
  styles?: StylesModel;
  creator?: string;
  lastModifiedBy?: string;
  lastPrinted?: Date;
  created?: Date;
  modified?: Date;
  company?: string;
  manager?: string;
  title?: string;
  subject?: string;
  keywords?: string;
  category?: string;
  description?: string;
  language?: string;
  revision?: number;
  contentStatus?: string;
  useSharedStrings?: boolean;
  useStyles?: boolean;
  pivotTables?: PivotTableModel[];
  commentRefs?: unknown[];
}

/**
 * Worksheet model during streaming load
 */
export interface WorksheetItemModel {
  sheetNo?: string | number;
  id?: number;
  name?: string;
  state?: string;
  relationships?: RelationshipModel[];
  drawing?: DrawingModel;
  rels?: RelationshipModel[];
  comments?: CommentModel[];
  tables?: TableModel[];
  [key: string]: unknown;
}

/**
 * Relationship model (internal Excel relationships)
 */
export interface RelationshipModel {
  Id: string;
  Type: string;
  Target: string;
  TargetMode?: string;
}

/**
 * Media/Image model
 */
export interface MediaModel {
  type: string;
  name: string;
  extension: string;
  buffer?: Uint8Array;
  filename?: string;
  base64?: string;
}

/**
 * Drawing model for charts, shapes, etc.
 */
export interface DrawingModel {
  name?: string;
  anchors?: AnchorModel[];
  rels?: RelationshipModel[];
  [key: string]: unknown;
}

/**
 * Anchor for drawings (positioning)
 */
export interface AnchorModel {
  picture?: {
    hyperlinks?: {
      rId?: string;
      hyperlink?: string;
    };
  };
  [key: string]: unknown;
}

/**
 * Comments model
 */
export interface CommentsModel {
  [key: string]: unknown;
}

/**
 * Individual comment
 */
export interface CommentModel {
  [key: string]: unknown;
}

/**
 * Table model
 */
export interface TableModel {
  target?: string;
  id?: number;
  name?: string;
  displayName?: string;
  [key: string]: unknown;
}

/**
 * VML Drawing model for legacy comments
 */
export interface VmlDrawingModel {
  [key: string]: unknown;
}

/**
 * Shared strings storage
 */
export interface SharedStringsModel {
  count?: number;
  xml?: string;
  [key: string]: unknown;
}

/**
 * Styles storage
 */
export interface StylesModel {
  xml?: string;
  [key: string]: unknown;
}

/**
 * Sheet definition from workbook
 */
export interface SheetDefinition {
  id: number | string;
  name: string;
  state?: string;
  [key: string]: unknown;
}

/**
 * Pivot table model
 */
export interface PivotTableModel {
  rId?: string;
  [key: string]: unknown;
}

/**
 * Worksheet model returned by WorksheetXform.parseStream()
 */
export interface ParsedWorksheetModel extends Omit<WorksheetItemModel, 'tables'> {
  dimensions?: unknown;
  cols?: unknown[];
  rows?: unknown[];
  mergeCells?: unknown;
  hyperlinks?: unknown[];
  dataValidations?: unknown[];
  properties?: Record<string, unknown>;
  views?: unknown[];
  pageSetup?: Record<string, unknown>;
  headerFooter?: unknown;
  background?: unknown;
  drawing?: DrawingModel;
  tables?: TableModel[] | unknown[];
  conditionalFormattings?: unknown[];
  autoFilter?: unknown;
  sheetProtection?: Record<string, unknown>;
  sheetNo?: string | number;
}

/**
 * Comments model returned by CommentsXform.parseStream()
 */
export interface ParsedCommentsModel extends CommentsModel {
  comments?: CommentModel[];
}

/**
 * Table model returned by TableXform.parseStream()
 */
export interface ParsedTableModel extends TableModel {
  tableRef?: string;
  totalsRow?: boolean;
  headerRow?: boolean;
}

/**
 * Relationships array returned by RelationshipsXform.parseStream()
 */
export type ParsedRelationshipsModel = RelationshipModel[];

/**
 * Drawing model returned by DrawingXform.parseStream()
 */
export interface ParsedDrawingModel extends DrawingModel {
  anchors: AnchorModel[];
}

/**
 * VML Drawing model returned by VmlNotesXform.parseStream()
 */
export interface ParsedVmlDrawingModel {
  [key: string]: unknown;
}

/**
 * Workbook model returned by WorkbookXform.parseStream()
 */
export interface ParsedWorkbookModel {
  sheets?: SheetDefinition[];
  definedNames?: unknown;
  views?: unknown[];
  properties?: Record<string, unknown>;
  calcProperties?: unknown;
}

/**
 * Shared strings model returned by SharedStringsXform.parseStream()
 */
export interface ParsedSharedStringsModel {
  [key: string]: unknown;
}

/**
 * App properties model returned by AppXform.parseStream()
 */
export interface ParsedAppPropertiesModel {
  company?: string;
  manager?: string;
  [key: string]: unknown;
}

/**
 * Core properties model returned by CoreXform.parseStream()
 */
export interface ParsedCorePropertiesModel {
  creator?: string;
  lastModifiedBy?: string;
  lastPrinted?: Date;
  created?: Date;
  modified?: Date;
  title?: string;
  subject?: string;
  keywords?: string;
  category?: string;
  description?: string;
  language?: string;
  revision?: number;
  contentStatus?: string;
  [key: string]: unknown;
}

/**
 * JSZip entry interface with stream-like capabilities
 */
export interface ZipEntry {
  name: string;
  dir: boolean;
  async(type: 'uint8array'): Promise<Uint8Array>;
  async(type: 'string'): Promise<string>;
  on?(event: 'error', handler: (error: Error) => void): void;
  pipe?(destination: WritableStream): void;
}

/**
 * ZipStream writer interface
 */
export interface ZipWriter {
  pipe(destination: XlsxWritableStream | unknown): ZipWriter;
  append(data: string | Uint8Array, options: { name: string; base64?: boolean }): Promise<void> | void;
  finalize(): Promise<void> | void;
  on(event: 'finish' | 'error', handler: (...args: unknown[]) => void): void;
}

/**
 * Custom array with add method for merge cells
 */
export interface MergesArray extends Array<unknown> {
  add?: () => void;
}

/**
 * Sheet options passed to StreamWriter classes
 */
export interface StreamWriterOptions {
  id: number;
  name?: string;
  state?: string;
  properties?: Record<string, unknown>;
  headerFooter?: Record<string, unknown>;
  pageSetup?: Record<string, unknown>;
  useSharedStrings?: boolean;
  workbook: unknown;
  views?: unknown[];
  autoFilter?: unknown;
  [key: string]: unknown;
}

/**
 * Workbook interface for streaming operations
 */
export interface StreamingWorkbook {
  _openStream(path: string): unknown;
  [key: string]: unknown;
}

/**
 * Stream interface with pause method
 */
export interface PausableStream extends WritableStream {
  pause?(): void;
  [key: string]: unknown;
}

/**
 * Relationship type constants
 */
export interface RelTypeRecord {
  OfficeDocument: string;
  Worksheet: string;
  CalcChain: string;
  SharedStrings: string;
  Styles: string;
  Theme: string;
  Hyperlink: string;
  Image: string;
  CoreProperties: string;
  ExtenderProperties: string;
  Comments: string;
  VmlDrawing: string;
  Table: string;
  PivotCacheDefinition: string;
  PivotCacheRecords: string;
  PivotTable: string;
}

/**
 * Transform xform interface (for XML transformers)
 */
export interface Xform {
  parseStream(stream: ReadableStream): Promise<unknown>;
  toXml(model: unknown): string;
  prepare(model: unknown, options?: unknown): void;
  render(xmlStream: unknown, model: unknown): void;
  reconcile(model: unknown, options?: unknown): void;
}
