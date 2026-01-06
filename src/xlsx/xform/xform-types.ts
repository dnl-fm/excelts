/**
 * Shared types for xform classes.
 * 
 * These types represent the data structures passed through the xform hierarchy
 * during prepare, render, and reconcile operations.
 */

import type Range from '../../doc/range.ts';

// =============================================================================
// XML Stream Writer
// =============================================================================

export type XmlStreamWriter = {
  openNode: (name: string, attributes?: Record<string, unknown>) => void;
  closeNode: (name?: string) => void;
  leafNode: (name: string, attributes: Record<string, unknown> | null, content?: unknown) => void;
  addAttribute: (name: string, value: unknown) => void;
  addAttributes: (attrs: Record<string, unknown>) => void;
  writeText: (text: string) => void;
  writeXml: (xml: string) => void;
  openXml: (attributes: Record<string, string>) => void;
  addRollback: () => void;
  rollback: () => void;
  commit: () => void;
};

// =============================================================================
// SAX Parser Node
// =============================================================================

export type SaxNode = {
  name: string;
  attributes: Record<string, string>;
};

// =============================================================================
// Style Manager
// =============================================================================

export type StylesManager = {
  addStyleModel: (style: unknown, effectiveCellType?: number) => number | undefined;
  getStyleModel: (styleId: number) => StyleModel | undefined;
  addDxfStyle: (style: unknown) => number;
  getDxfStyle: (dxfId: number) => unknown;
};

export type StyleModel = {
  numFmt?: string;
  [key: string]: unknown;
};

// =============================================================================
// Shared Strings Manager
// =============================================================================

export type SharedStringsManager = {
  add: (value: unknown) => number;
  getString: (index: number) => unknown;
};

// =============================================================================
// Cell Types
// =============================================================================

export type CellModel = {
  address?: string;
  type?: number;
  value?: unknown;
  text?: string;
  result?: unknown;
  styleId?: number;
  style?: unknown;
  ssId?: number;
  formula?: string;
  /** 
   * Shared formula index. 
   * During parsing this is a string from XML attributes.
   * During prepare/reconcile it may be a number.
   */
  si?: number | string;
  shareType?: string;
  ref?: string;
  range?: Range;
  sharedFormula?: string;
  date1904?: boolean;
  hyperlink?: string;
  tooltip?: string;
  comment?: unknown;
};

export type HyperlinkRef = {
  address?: string;
  target?: string;
  tooltip?: string;
  rId?: string;
};

export type CommentRef = {
  ref?: string;
  [key: string]: unknown;
};

export type MergesManager = {
  add: (model: CellModel) => void;
  mergeCells: unknown[];
};

// =============================================================================
// Cell Xform Options
// =============================================================================

/**
 * Options passed to CellXform.prepare()
 * 
 * Built up by worksheet-xform.prepare() and passed down through:
 * worksheet-xform -> ListXform (sheetData) -> row-xform -> cell-xform
 * 
 * Note: This object is mutated during prepare (siFormulae is incremented,
 * arrays are pushed to). This is the existing design pattern.
 */
export type CellXformPrepareOptions = {
  styles: StylesManager;
  comments: CommentRef[];
  sharedStrings?: SharedStringsManager;
  date1904?: boolean;
  hyperlinks: HyperlinkRef[];
  merges: MergesManager;
  formulae: Record<string, CellModel>;
  siFormulae: number;
};

/**
 * Options passed to CellXform.reconcile()
 * 
 * Used during post-parse processing to resolve references.
 */
export type CellXformReconcileOptions = {
  styles?: StylesManager;
  sharedStrings?: SharedStringsManager;
  date1904?: boolean;
  formulae: Record<number, string>;
  hyperlinkMap: Record<string, string>;
  commentsMap?: Record<string, unknown>;
};

// =============================================================================
// Row Types
// =============================================================================

export type RowModel = {
  number: number;
  min?: number;
  max?: number;
  cells: CellModel[];
  styleId?: number;
  style?: unknown;
  hidden?: boolean;
  bestFit?: boolean;
  height?: number;
  outlineLevel?: number;
  collapsed?: boolean;
  [key: string]: unknown;
};

/**
 * Options passed to RowXform.prepare() and RowXform.reconcile()
 * 
 * Same as CellXformPrepareOptions since it's passed through.
 */
export type RowXformPrepareOptions = CellXformPrepareOptions;
export type RowXformReconcileOptions = CellXformReconcileOptions;
