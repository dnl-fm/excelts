/**
 * Core types for ExcelTS XLSX I/O operations
 */

import { Readable, Writable } from 'readable-stream';
import type JSZip from 'jszip';
import type ZipStream from '../utils/zip-stream.ts';

/**
 * Stream types - unified interface for Node.js and browser streams
 */
export type ReadableStream = Readable | AsyncIterable<Buffer | string>;
export type WritableStream = Writable;

/**
 * XLSX options for read operations
 */
export interface XlsxReadOptions {
  base64?: boolean;
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
  buffer?: Buffer;
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
 * JSZip entry interface
 */
export interface ZipEntry {
  name: string;
  dir: boolean;
  async(type: 'nodebuffer'): Promise<Buffer>;
  async(type: 'string'): Promise<string>;
}

/**
 * ZipStream writer interface
 */
export interface ZipWriter {
  pipe(destination: WritableStream): ZipWriter;
  append(data: string | Buffer, options: { name: string; base64?: boolean }): Promise<void> | void;
  finalize(): Promise<void> | void;
  on(event: 'finish' | 'error', handler: (...args: unknown[]) => void): void;
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
