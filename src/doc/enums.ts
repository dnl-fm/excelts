export const ValueType = {
  Null: 0,
  Merge: 1,
  Number: 2,
  String: 3,
  Date: 4,
  Hyperlink: 5,
  Formula: 6,
  SharedString: 7,
  RichText: 8,
  Boolean: 9,
  Error: 10,
  JSON: 11,
};

export const FormulaType = {
  None: 0,
  Master: 1,
  Shared: 2,
};

export const RelationshipType = {
  None: 0,
  OfficeDocument: 1,
  Worksheet: 2,
  CalcChain: 3,
  SharedStrings: 4,
  Styles: 5,
  Theme: 6,
  Hyperlink: 7,
};

export const DocumentType = {
  Xlsx: 1,
};

export const ReadingOrder = {
  LeftToRight: 1,
  RightToLeft: 2,
};

export const ErrorValue = {
  NotApplicable: '#N/A',
  Ref: '#REF!',
  Name: '#NAME?',
  DivZero: '#DIV/0!',
  Null: '#NULL!',
  Value: '#VALUE!',
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
