/* eslint-disable max-classes-per-file */
import colCache from '../utils/col-cache.ts';
import _ from '../utils/under-dash.ts';
import Enums from './enums.ts';
import Note from './note.ts';
import { slideFormula } from '../utils/shared-formula.ts';
import type Row from './row.ts';
import type Column from './column.ts';
import type Worksheet from './worksheet.ts';
import type Workbook from './workbook.ts';

/**
 * Represents a single cell in a worksheet.
 *
 * Cells store values, formulas, and styles. Values can be numbers, strings,
 * dates, booleans, rich text, hyperlinks, or formulas.
 *
 * @example
 * ```ts
 * const cell = sheet.getCell('A1');
 *
 * // Set value
 * cell.value = 'Hello';
 * cell.value = 42;
 * cell.value = new Date();
 * cell.value = { formula: 'SUM(B1:B10)' };
 *
 * // Set style
 * cell.font = { bold: true, size: 14 };
 * cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
 * cell.border = { top: { style: 'thin' } };
 * cell.alignment = { horizontal: 'center' };
 * ```
 */
class Cell {
  /** @internal */
  _row: Row;
  /** @internal */
  _column: Column;
  /** @internal */
  _address: string;
  /** @internal */
  _value: unknown;
  /** @internal */
  style: Record<string, unknown>;
  /** @internal */
  _mergeCount: number;

  constructor(row: Row, column: Column, address: string) {
    if (!row || !column) {
      throw new Error('A Cell needs a Row');
    }

    this._row = row;
    this._column = column;

    colCache.validateAddress(address);
    this._address = address;

    this._value = Value.create(Cell.Types.Null, this);
    this.style = this._mergeStyle(row.style, column.style, {});
    this._mergeCount = 0;
  }

  get worksheet(): Worksheet {
    return this._row.worksheet;
  }

  get workbook(): Workbook {
    return this._row.worksheet.workbook;
  }

  // help GC by removing cyclic (and other) references
  destroy(): void {
    delete this.style;
    delete this._value;
    delete this._row;
    delete this._column;
    delete this._address;
  }

  // =========================================================================
  // Styles

  /**
   * Number format string (e.g., '0.00', '#,##0', 'yyyy-mm-dd').
   *
   * @example
   * ```ts
   * cell.numFmt = '0.00%';      // percentage
   * cell.numFmt = '$#,##0.00';  // currency
   * cell.numFmt = 'yyyy-mm-dd'; // date
   * ```
   */
  get numFmt(): unknown {
    return this.style.numFmt;
  }

  set numFmt(value: unknown) {
    this.style.numFmt = value;
  }

  /**
   * Font styling for the cell.
   *
   * @example
   * ```ts
   * cell.font = {
   *   name: 'Arial',
   *   size: 12,
   *   bold: true,
   *   italic: false,
   *   underline: true,
   *   color: { argb: 'FF0000FF' },
   * };
   * ```
   */
  get font(): unknown {
    return this.style.font;
  }

  set font(value: unknown) {
    this.style.font = value;
  }

  /**
   * Text alignment within the cell.
   *
   * @example
   * ```ts
   * cell.alignment = {
   *   horizontal: 'center',  // left, center, right, fill, justify
   *   vertical: 'middle',    // top, middle, bottom
   *   wrapText: true,
   *   textRotation: 45,
   * };
   * ```
   */
  get alignment(): unknown {
    return this.style.alignment;
  }

  set alignment(value: unknown) {
    this.style.alignment = value;
  }

  /**
   * Cell border styling.
   *
   * @example
   * ```ts
   * cell.border = {
   *   top: { style: 'thin', color: { argb: 'FF000000' } },
   *   left: { style: 'thin' },
   *   bottom: { style: 'double' },
   *   right: { style: 'thin' },
   * };
   * ```
   */
  get border(): unknown {
    return this.style.border;
  }

  set border(value: unknown) {
    this.style.border = value;
  }

  /**
   * Cell fill (background) styling.
   *
   * @example
   * ```ts
   * // Solid fill
   * cell.fill = {
   *   type: 'pattern',
   *   pattern: 'solid',
   *   fgColor: { argb: 'FFFFFF00' },
   * };
   *
   * // Gradient fill
   * cell.fill = {
   *   type: 'gradient',
   *   gradient: 'angle',
   *   degree: 90,
   *   stops: [
   *     { position: 0, color: { argb: 'FF0000FF' } },
   *     { position: 1, color: { argb: 'FF00FF00' } },
   *   ],
   * };
   * ```
   */
  get fill(): unknown {
    return this.style.fill;
  }

  set fill(value: unknown) {
    this.style.fill = value;
  }

  /**
   * Cell protection settings (used with sheet protection).
   *
   * @example
   * ```ts
   * cell.protection = { locked: false };
   * ```
   */
  get protection(): unknown {
    return this.style.protection;
  }

  set protection(value: unknown) {
    this.style.protection = value;
  }

  _mergeStyle(rowStyle: Record<string, unknown>, colStyle: Record<string, unknown>, style: Record<string, unknown>): Record<string, unknown> {
    const numFmt = (rowStyle && rowStyle.numFmt) || (colStyle && colStyle.numFmt);
    if (numFmt) style.numFmt = numFmt;

    const font = (rowStyle && rowStyle.font) || (colStyle && colStyle.font);
    if (font) style.font = font;

    const alignment = (rowStyle && rowStyle.alignment) || (colStyle && colStyle.alignment);
    if (alignment) style.alignment = alignment;

    const border = (rowStyle && rowStyle.border) || (colStyle && colStyle.border);
    if (border) style.border = border;

    const fill = (rowStyle && rowStyle.fill) || (colStyle && colStyle.fill);
    if (fill) style.fill = fill;

    const protection = (rowStyle && rowStyle.protection) || (colStyle && colStyle.protection);
    if (protection) style.protection = protection;

    return style;
  }

  // =========================================================================
  // Address

  /** Cell address (e.g., 'A1', 'B2') */
  get address(): string {
    return this._address;
  }

  /** Row number (1-based) */
  get row(): number {
    return this._row.number;
  }

  /** Column number (1-based) */
  get col(): number {
    return this._column.number;
  }

  /** Absolute address (e.g., '$A$1') */
  get $col$row(): string {
    return `$${this._column.letter}$${this.row}`;
  }

  // =========================================================================
  // Value

  /**
   * The cell's value type (Null, Number, String, Date, etc.).
   * See ValueType enum for all types.
   */
  get type(): string {
    return this._value.type;
  }

  /**
   * The effective value type (resolves formula results).
   */
  get effectiveType(): string {
    return this._value.effectiveType;
  }

  /**
   * Returns the cell value formatted for CSV export.
   */
  toCsvString(): string {
    return this._value.toCsvString();
  }

  // =========================================================================
  // Merge stuff

  addMergeRef(): void {
    this._mergeCount++;
  }

  releaseMergeRef(): void {
    this._mergeCount--;
  }

  get isMerged(): boolean {
    return this._mergeCount > 0 || this.type === Cell.Types.Merge;
  }

  merge(master: Cell, ignoreStyle?: boolean): void {
    this._value.release();
    this._value = Value.create(Cell.Types.Merge, this, master);
    if (!ignoreStyle) {
      this.style = master.style;
    }
  }

  unmerge(): void {
    if (this.type === Cell.Types.Merge) {
      this._value.release();
      this._value = Value.create(Cell.Types.Null, this);
      this.style = this._mergeStyle(this._row.style, this._column.style, {});
    }
  }

  isMergedTo(master: Cell): boolean {
    if (this._value.type !== Cell.Types.Merge) return false;
    return this._value.isMergedTo(master);
  }

  /**
   * For merged cells, returns the master cell. Unmerged cells return themselves.
   */
  get master(): Cell {
    if (this.type === Cell.Types.Merge) {
      return this._value.master;
    }
    return this; // an unmerged cell is its own master
  }

  /** Whether this cell contains a hyperlink */
  get isHyperlink(): boolean {
    return this._value.type === Cell.Types.Hyperlink;
  }

  /** The hyperlink URL (if cell is a hyperlink) */
  get hyperlink(): unknown {
    return this._value.hyperlink;
  }

  /**
   * The cell's value.
   *
   * @example
   * ```ts
   * // Get value
   * const val = cell.value;
   *
   * // Set simple values
   * cell.value = 'Hello';
   * cell.value = 42;
   * cell.value = new Date();
   * cell.value = true;
   *
   * // Set formula
   * cell.value = { formula: 'SUM(A1:A10)', result: 55 };
   *
   * // Set hyperlink
   * cell.value = { text: 'Click here', hyperlink: 'https://example.com' };
   *
   * // Set rich text
   * cell.value = {
   *   richText: [
   *     { text: 'Bold ', font: { bold: true } },
   *     { text: 'Normal' },
   *   ],
   * };
   * ```
   */
  get value(): unknown {
    return this._value.value;
  }

  set value(v: unknown) {
    // special case - merge cells set their master's value
    if (this.type === Cell.Types.Merge) {
      this._value.master.value = v;
      return;
    }

    this._value.release();

    // assign value
    this._value = Value.create(Value.getType(v), this, v);
  }

  /**
   * Cell comment/note.
   *
   * @example
   * ```ts
   * cell.note = 'This is a comment';
   * cell.note = {
   *   texts: [{ text: 'Comment text' }],
   *   margins: { left: 0.1, right: 0.1, top: 0.1, bottom: 0.1 },
   * };
   * ```
   */
  get note(): unknown {
    return this._comment && this._comment.note;
  }

  set note(note: unknown) {
    this._comment = new Note(note);
  }

  /** The cell value as a plain text string */
  get text(): string {
    return this._value.toString();
  }

  /** The cell value as HTML-escaped text */
  get html(): string {
    return _.escapeHtml(this.text);
  }

  toString(): string {
    return this.text;
  }

  _upgradeToHyperlink(hyperlink: string): void {
    // if this cell is a string, turn it into a Hyperlink
    if (this.type === Cell.Types.String) {
      this._value = Value.create(Cell.Types.Hyperlink, this, {
        text: this._value.value,
        hyperlink,
      });
    }
  }

  // =========================================================================
  // Formula

  /**
   * The cell's formula (without the leading '=').
   * Returns undefined for non-formula cells.
   */
  get formula(): string | undefined {
    return this._value.formula;
  }

  /**
   * The calculated result of the formula.
   * Returns undefined for non-formula cells.
   */
  get result(): unknown {
    return this._value.result;
  }

  /**
   * The formula type: Master, Shared, or None.
   */
  get formulaType(): string | undefined {
    return this._value.formulaType;
  }

  // =========================================================================
  // Name stuff
  get fullAddress(): Record<string, unknown> {
    const {worksheet} = this._row;
    return {
      sheetName: worksheet.name,
      address: this.address,
      row: this.row,
      col: this.col,
    };
  }

  get name(): unknown {
    return this.names[0];
  }

  set name(value: unknown) {
    this.names = [value];
  }

  get names(): unknown[] {
    return this.workbook.definedNames.getNamesEx(this.fullAddress);
  }

  set names(value: unknown[]) {
    const {definedNames} = this.workbook;
    definedNames.removeAllNames(this.fullAddress);
    value.forEach((name: unknown) => {
      definedNames.addEx(this.fullAddress, name);
    });
  }

  addName(name: unknown): void {
    this.workbook.definedNames.addEx(this.fullAddress, name);
  }

  removeName(name: unknown): void {
    this.workbook.definedNames.removeEx(this.fullAddress, name);
  }

  removeAllNames(): void {
    this.workbook.definedNames.removeAllNames(this.fullAddress);
  }

  // =========================================================================
  // Data Validation stuff
  get _dataValidations(): unknown {
    return (this.worksheet as Record<string, unknown>).dataValidations;
  }

  get dataValidation(): unknown {
    const dv = this._dataValidations as Record<string, unknown>;
    return (dv.find as Function)(this.address);
  }

  set dataValidation(value: unknown) {
    const dv = this._dataValidations as Record<string, unknown>;
    (dv.add as Function)(this.address, value);
  }

  // =========================================================================
  // Model stuff

  get model(): Record<string, unknown> {
    const {model} = this._value;
    model.style = this.style;
    if (this._comment) {
      model.comment = this._comment.model;
    }
    return model;
  }

  set model(value: Record<string, unknown>) {
    const oldValue = this._value as Record<string, unknown>;
    oldValue.release?.();
    // Use Cell.Types.Null (0) as default when type is null/undefined
    const cellType = value.type ?? Cell.Types.Null;
    this._value = Value.create(cellType, this);
    const newValue = this._value as Record<string, unknown>;
    newValue.model = value;

    if (value.comment) {
      const comment = value.comment as Record<string, unknown>;
      switch (comment.type) {
        case 'note':
          this._comment = Note.fromModel(comment);
          break;
      }
    }

    if (value.style) {
      this.style = value.style;
    } else {
      this.style = {};
    }
  }
}
Cell.Types = Enums.ValueType;

// =============================================================================
// Internal Value Types

class NullValue {
  constructor(cell) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Null,
    };
  }

  get value() {
    return null;
  }

  set value(value) {
    // nothing to do
  }

  get type() {
    return Cell.Types.Null;
  }

  get effectiveType() {
    return Cell.Types.Null;
  }

  get address() {
    return this.model.address;
  }

  set address(value) {
    this.model.address = value;
  }

  toCsvString() {
    return '';
  }

  release() {}

  toString() {
    return '';
  }
}

class NumberValue {
  constructor(cell, value) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Number,
      value,
    };
  }

  get value() {
    return this.model.value;
  }

  set value(value) {
    this.model.value = value;
  }

  get type() {
    return Cell.Types.Number;
  }

  get effectiveType() {
    return Cell.Types.Number;
  }

  get address() {
    return this.model.address;
  }

  set address(value) {
    this.model.address = value;
  }

  toCsvString() {
    return this.model.value.toString();
  }

  release() {}

  toString() {
    return this.model.value.toString();
  }
}

class StringValue {
  constructor(cell, value) {
    this.model = {
      address: cell.address,
      type: Cell.Types.String,
      value,
    };
  }

  get value() {
    return this.model.value;
  }

  set value(value) {
    this.model.value = value;
  }

  get type() {
    return Cell.Types.String;
  }

  get effectiveType() {
    return Cell.Types.String;
  }

  get address() {
    return this.model.address;
  }

  set address(value) {
    this.model.address = value;
  }

  toCsvString() {
    return `"${this.model.value.replace(/"/g, '""')}"`;
  }

  release() {}

  toString() {
    return this.model.value;
  }
}

class RichTextValue {
  constructor(cell, value) {
    this.model = {
      address: cell.address,
      type: Cell.Types.String,
      value,
    };
  }

  get value() {
    return this.model.value;
  }

  set value(value) {
    this.model.value = value;
  }

  toString() {
    return this.model.value.richText.map(t => t.text).join('');
  }

  get type() {
    return Cell.Types.RichText;
  }

  get effectiveType() {
    return Cell.Types.RichText;
  }

  get address() {
    return this.model.address;
  }

  set address(value) {
    this.model.address = value;
  }

  toCsvString() {
    return `"${this.text.replace(/"/g, '""')}"`;
  }

  release() {}
}

class DateValue {
  constructor(cell, value) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Date,
      value,
    };
  }

  get value() {
    return this.model.value;
  }

  set value(value) {
    this.model.value = value;
  }

  get type() {
    return Cell.Types.Date;
  }

  get effectiveType() {
    return Cell.Types.Date;
  }

  get address() {
    return this.model.address;
  }

  set address(value) {
    this.model.address = value;
  }

  toCsvString() {
    return this.model.value.toISOString();
  }

  release() {}

  toString() {
    return this.model.value.toString();
  }
}

class HyperlinkValue {
  constructor(cell, value) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Hyperlink,
      text: value ? value.text : undefined,
      hyperlink: value ? value.hyperlink : undefined,
    };
    if (value && value.tooltip) {
      this.model.tooltip = value.tooltip;
    }
  }

  get value() {
    const v = {
      text: this.model.text,
      hyperlink: this.model.hyperlink,
    };
    if (this.model.tooltip) {
      v.tooltip = this.model.tooltip;
    }
    return v;
  }

  set value(value) {
    this.model = {
      text: value.text,
      hyperlink: value.hyperlink,
    };
    if (value.tooltip) {
      this.model.tooltip = value.tooltip;
    }
  }

  get text() {
    return this.model.text;
  }

  set text(value) {
    this.model.text = value;
  }

  /*
  get tooltip() {
    return this.model.tooltip;
  }

  set tooltip(value) {
    this.model.tooltip = value;
  } */

  get hyperlink() {
    return this.model.hyperlink;
  }

  set hyperlink(value) {
    this.model.hyperlink = value;
  }

  get type() {
    return Cell.Types.Hyperlink;
  }

  get effectiveType() {
    return Cell.Types.Hyperlink;
  }

  get address() {
    return this.model.address;
  }

  set address(value) {
    this.model.address = value;
  }

  toCsvString() {
    return this.model.hyperlink;
  }

  release() {}

  toString() {
    return this.model.text;
  }
}

class MergeValue {
  constructor(cell, master) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Merge,
      master: master ? master.address : undefined,
    };
    this._master = master;
    if (master) {
      master.addMergeRef();
    }
  }

  get value() {
    return this._master.value;
  }

  set value(value) {
    if (value instanceof Cell) {
      if (this._master) {
        this._master.releaseMergeRef();
      }
      value.addMergeRef();
      this._master = value;
    } else {
      this._master.value = value;
    }
  }

  isMergedTo(master) {
    return master === this._master;
  }

  get master() {
    return this._master;
  }

  get type() {
    return Cell.Types.Merge;
  }

  get effectiveType() {
    return this._master.effectiveType;
  }

  get address() {
    return this.model.address;
  }

  set address(value) {
    this.model.address = value;
  }

  toCsvString() {
    return '';
  }

  release() {
    this._master.releaseMergeRef();
  }

  toString() {
    return this.value.toString();
  }
}

class FormulaValue {
  constructor(cell, value) {
    this.cell = cell;

    this.model = {
      address: cell.address,
      type: Cell.Types.Formula,
      shareType: value ? value.shareType : undefined,
      ref: value ? value.ref : undefined,
      formula: value ? value.formula : undefined,
      sharedFormula: value ? value.sharedFormula : undefined,
      result: value ? value.result : undefined,
    };
  }

  _copyModel(model) {
    const copy = {};
    const cp = name => {
      const value = model[name];
      if (value) {
        copy[name] = value;
      }
    };
    cp('formula');
    cp('result');
    cp('ref');
    cp('shareType');
    cp('sharedFormula');
    return copy;
  }

  get value() {
    return this._copyModel(this.model);
  }

  set value(value) {
    this.model = this._copyModel(value);
  }

  validate(value) {
    switch (Value.getType(value)) {
      case Cell.Types.Null:
      case Cell.Types.String:
      case Cell.Types.Number:
      case Cell.Types.Date:
        break;
      case Cell.Types.Hyperlink:
      case Cell.Types.Formula:
      default:
        throw new Error('Cannot process that type of result value');
    }
  }

  get dependencies() {
    // find all the ranges and cells mentioned in the formula
    const ranges = this.formula.match(/([a-zA-Z0-9]+!)?[A-Z]{1,3}\d{1,4}:[A-Z]{1,3}\d{1,4}/g);
    const cells = this.formula
      .replace(/([a-zA-Z0-9]+!)?[A-Z]{1,3}\d{1,4}:[A-Z]{1,3}\d{1,4}/g, '')
      .match(/([a-zA-Z0-9]+!)?[A-Z]{1,3}\d{1,4}/g);
    return {
      ranges,
      cells,
    };
  }

  get formula() {
    return this.model.formula || this._getTranslatedFormula();
  }

  set formula(value) {
    this.model.formula = value;
  }

  get formulaType() {
    if (this.model.formula) {
      return Enums.FormulaType.Master;
    }
    if (this.model.sharedFormula) {
      return Enums.FormulaType.Shared;
    }
    return Enums.FormulaType.None;
  }

  get result() {
    return this.model.result;
  }

  set result(value) {
    this.model.result = value;
  }

  get type() {
    return Cell.Types.Formula;
  }

  get effectiveType() {
    const v = this.model.result;
    if (v === null || v === undefined) {
      return Enums.ValueType.Null;
    }
    if (v instanceof String || typeof v === 'string') {
      return Enums.ValueType.String;
    }
    if (typeof v === 'number') {
      return Enums.ValueType.Number;
    }
    if (v instanceof Date) {
      return Enums.ValueType.Date;
    }
    if (v.text && v.hyperlink) {
      return Enums.ValueType.Hyperlink;
    }
    if (v.formula) {
      return Enums.ValueType.Formula;
    }

    return Enums.ValueType.Null;
  }

  get address() {
    return this.model.address;
  }

  set address(value) {
    this.model.address = value;
  }

  _getTranslatedFormula() {
    if (!this._translatedFormula && this.model.sharedFormula) {
      const {worksheet} = this.cell;
      const master = worksheet.findCell(this.model.sharedFormula);
      this._translatedFormula =
        master && slideFormula(master.formula, master.address, this.model.address);
    }
    return this._translatedFormula;
  }

  toCsvString() {
    return `${this.model.result || ''}`;
  }

  release() {}

  toString() {
    return this.model.result ? this.model.result.toString() : '';
  }
}

class SharedStringValue {
  constructor(cell, value) {
    this.model = {
      address: cell.address,
      type: Cell.Types.SharedString,
      value,
    };
  }

  get value() {
    return this.model.value;
  }

  set value(value) {
    this.model.value = value;
  }

  get type() {
    return Cell.Types.SharedString;
  }

  get effectiveType() {
    return Cell.Types.SharedString;
  }

  get address() {
    return this.model.address;
  }

  set address(value) {
    this.model.address = value;
  }

  toCsvString() {
    return this.model.value.toString();
  }

  release() {}

  toString() {
    return this.model.value.toString();
  }
}

class BooleanValue {
  constructor(cell, value) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Boolean,
      value,
    };
  }

  get value() {
    return this.model.value;
  }

  set value(value) {
    this.model.value = value;
  }

  get type() {
    return Cell.Types.Boolean;
  }

  get effectiveType() {
    return Cell.Types.Boolean;
  }

  get address() {
    return this.model.address;
  }

  set address(value) {
    this.model.address = value;
  }

  toCsvString() {
    return this.model.value ? 1 : 0;
  }

  release() {}

  toString() {
    return this.model.value.toString();
  }
}

class ErrorValue {
  constructor(cell, value) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Error,
      value,
    };
  }

  get value() {
    return this.model.value;
  }

  set value(value) {
    this.model.value = value;
  }

  get type() {
    return Cell.Types.Error;
  }

  get effectiveType() {
    return Cell.Types.Error;
  }

  get address() {
    return this.model.address;
  }

  set address(value) {
    this.model.address = value;
  }

  toCsvString() {
    return this.toString();
  }

  release() {}

  toString() {
    return this.model.value.error.toString();
  }
}

class JSONValue {
  constructor(cell, value) {
    this.model = {
      address: cell.address,
      type: Cell.Types.String,
      value: JSON.stringify(value),
      rawValue: value,
    };
  }

  get value() {
    return this.model.rawValue;
  }

  set value(value) {
    this.model.rawValue = value;
    this.model.value = JSON.stringify(value);
  }

  get type() {
    return Cell.Types.String;
  }

  get effectiveType() {
    return Cell.Types.String;
  }

  get address() {
    return this.model.address;
  }

  set address(value) {
    this.model.address = value;
  }

  toCsvString() {
    return this.model.value;
  }

  release() {}

  toString() {
    return this.model.value;
  }
}

// Value is a place to hold common static Value type functions
const Value = {
  getType(value) {
    if (value === null || value === undefined) {
      return Cell.Types.Null;
    }
    if (value instanceof String || typeof value === 'string') {
      return Cell.Types.String;
    }
    if (typeof value === 'number') {
      return Cell.Types.Number;
    }
    if (typeof value === 'boolean') {
      return Cell.Types.Boolean;
    }
    if (value instanceof Date) {
      return Cell.Types.Date;
    }
    if (value.text && value.hyperlink) {
      return Cell.Types.Hyperlink;
    }
    if (value.formula || value.sharedFormula) {
      return Cell.Types.Formula;
    }
    if (value.richText) {
      return Cell.Types.RichText;
    }
    if (value.sharedString) {
      return Cell.Types.SharedString;
    }
    if (value.error) {
      return Cell.Types.Error;
    }
    return Cell.Types.JSON;
  },

  // map valueType to constructor
  types: [
    {t: Cell.Types.Null, f: NullValue},
    {t: Cell.Types.Number, f: NumberValue},
    {t: Cell.Types.String, f: StringValue},
    {t: Cell.Types.Date, f: DateValue},
    {t: Cell.Types.Hyperlink, f: HyperlinkValue},
    {t: Cell.Types.Formula, f: FormulaValue},
    {t: Cell.Types.Merge, f: MergeValue},
    {t: Cell.Types.JSON, f: JSONValue},
    {t: Cell.Types.SharedString, f: SharedStringValue},
    {t: Cell.Types.RichText, f: RichTextValue},
    {t: Cell.Types.Boolean, f: BooleanValue},
    {t: Cell.Types.Error, f: ErrorValue},
  ].reduce((p, t) => {
    p[t.t] = t.f;
    return p;
  }, []),

  create(type, cell, value) {
    const T = this.types[type];
    if (!T) {
      throw new Error(`Could not create Value of type ${type}`);
    }
    return new T(cell, value);
  },
};

export default Cell;