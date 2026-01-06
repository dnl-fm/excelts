/**
 * Cell value type enumeration.
 * Identifies the type of data stored in a cell.
 *
 * @example
 * ```ts
 * if (cell.type === ValueType.Formula) {
 *   console.log('Formula:', cell.formula);
 * }
 * ```
 */
export const ValueType = {
  /** Empty cell with no value */
  Null: 0,
  /** Merged cell (references master cell) */
  Merge: 1,
  /** Numeric value */
  Number: 2,
  /** String/text value */
  String: 3,
  /** Date/time value */
  Date: 4,
  /** Hyperlink with text and URL */
  Hyperlink: 5,
  /** Formula (may have cached result) */
  Formula: 6,
  /** Shared string reference (internal) */
  SharedString: 7,
  /** Rich text with formatting */
  RichText: 8,
  /** Boolean true/false */
  Boolean: 9,
  /** Error value (#N/A, #REF!, etc.) */
  Error: 10,
  /** JSON object value */
  JSON: 11,
};

/**
 * Formula type enumeration.
 * Identifies how a formula is stored in the workbook.
 */
export const FormulaType = {
  /** Not a formula */
  None: 0,
  /** Master formula (original definition) */
  Master: 1,
  /** Shared formula (references master) */
  Shared: 2,
};

/**
 * Relationship type enumeration.
 * Used internally for Excel part relationships.
 */
export const RelationshipType = {
  /** No relationship */
  None: 0,
  /** Office document */
  OfficeDocument: 1,
  /** Worksheet */
  Worksheet: 2,
  /** Calculation chain */
  CalcChain: 3,
  /** Shared strings table */
  SharedStrings: 4,
  /** Styles */
  Styles: 5,
  /** Theme */
  Theme: 6,
  /** Hyperlink */
  Hyperlink: 7,
};

/**
 * Document type enumeration.
 */
export const DocumentType = {
  /** XLSX format */
  Xlsx: 1,
};

/**
 * Text reading order enumeration.
 * Used for cell alignment in bidirectional text.
 */
export const ReadingOrder = {
  /** Left to right (default for most languages) */
  LeftToRight: 1,
  /** Right to left (for RTL languages like Arabic, Hebrew) */
  RightToLeft: 2,
};

/**
 * Excel error value constants.
 * Standard error values that can appear in cells.
 *
 * @example
 * ```ts
 * cell.value = { error: ErrorValue.NotApplicable };
 * ```
 */
export const ErrorValue = {
  /** #N/A - Value not available */
  NotApplicable: '#N/A',
  /** #REF! - Invalid cell reference */
  Ref: '#REF!',
  /** #NAME? - Unrecognized formula name */
  Name: '#NAME?',
  /** #DIV/0! - Division by zero */
  DivZero: '#DIV/0!',
  /** #NULL! - Incorrect range operator */
  Null: '#NULL!',
  /** #VALUE! - Wrong type of argument */
  Value: '#VALUE!',
  /** #NUM! - Invalid numeric value */
  Num: '#NUM!',
};

// Default export for backward compatibility
export default {
  ValueType,
  FormulaType,
  RelationshipType,
  DocumentType,
  ReadingOrder,
  ErrorValue,
};
