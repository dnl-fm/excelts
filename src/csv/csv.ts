
import { inferSchema, initParser } from 'udsv';
import dayjs from 'dayjs';
import customParseFormat from 'dayjs/plugin/customParseFormat';
import utc from 'dayjs/plugin/utc';
import type Workbook from '../doc/workbook.ts';
import type Worksheet from '../doc/worksheet.ts';

dayjs.extend(customParseFormat);
dayjs.extend(utc);

/* eslint-disable quote-props */
const SpecialValues: Record<string, unknown> = {
  true: true,
  false: false,
  '#N/A': {error: '#N/A'},
  '#REF!': {error: '#REF!'},
  '#NAME?': {error: '#NAME?'},
  '#DIV/0!': {error: '#DIV/0!'},
  '#NULL!': {error: '#NULL!'},
  '#VALUE!': {error: '#VALUE!'},
  '#NUM!': {error: '#NUM!'},
};
/* eslint-enable quote-props */

interface CSVReadOptions {
  sheetName?: string;
  dateFormats?: string[];
  map?: (datum: unknown) => unknown;
  parserOptions?: {
    delimiter?: string;
    quote?: string;
    escape?: string;
  };
}

interface CSVWriteOptions {
  sheetName?: string;
  sheetId?: number;
  dateFormat?: string;
  dateUTC?: boolean;
  map?: (value: unknown) => unknown;
  includeEmptyRows?: boolean;
  formatterOptions?: {
    delimiter?: string;
    quote?: string;
    escape?: string;
    rowDelimiter?: string;
  };
}

interface CSVFileWriteOptions extends CSVWriteOptions {
  encoding?: BufferEncoding;
}

/**
 * Simple CSV formatter (replaces fast-csv's format)
 */
function formatCSVValue(value: unknown, quote = '"', escape = '"'): string {
  if (value === null || value === undefined) {
    return '';
  }
  const str = String(value);
  // Check if quoting is needed
  if (str.includes(quote) || str.includes(',') || str.includes('\n') || str.includes('\r')) {
    // Escape quotes and wrap in quotes
    return quote + str.replace(new RegExp(quote, 'g'), escape + quote) + quote;
  }
  return str;
}

function formatCSVRow(
  row: unknown[],
  delimiter = ',',
  quote = '"',
  escape = '"'
): string {
  return row.map(v => formatCSVValue(v, quote, escape)).join(delimiter);
}

/**
 * CSV handles reading and writing worksheets to CSV streams and files.
 */
class CSV {
  workbook: Workbook;
  worksheet: Worksheet | null;

  constructor(workbook: Workbook) {
    this.workbook = workbook;
    this.worksheet = null;
  }

  async readFile(filename: string, options?: CSVReadOptions): Promise<Worksheet> {
    const opts = options || {};
    const file = Bun.file(filename);
    if (!(await file.exists())) {
      throw new Error(`File not found: ${filename}`);
    }
    const content = await file.text();
    return this.parseString(content, opts);
  }

  /**
   * Read from an async iterable stream
   */
  async read(stream: AsyncIterable<string | Buffer>, options?: CSVReadOptions): Promise<Worksheet> {
    const chunks: string[] = [];
    for await (const chunk of stream) {
      chunks.push(typeof chunk === 'string' ? chunk : chunk.toString());
    }
    return this.parseString(chunks.join(''), options);
  }

  /**
   * Parse CSV string directly
   */
  parseString(csvString: string, options?: CSVReadOptions): Worksheet {
    const opts = options || {};
    const worksheet = this.workbook.addWorksheet(opts.sheetName);

    const dateFormats = opts.dateFormats || [
      'YYYY-MM-DD[T]HH:mm:ssZ',
      'YYYY-MM-DD[T]HH:mm:ss',
      'MM-DD-YYYY',
      'YYYY-MM-DD',
    ];

    const map =
      opts.map ||
      function(datum: unknown): unknown {
        if (datum === '' || datum === null || datum === undefined) {
          return null;
        }
        const datumNumber = Number(datum);
        if (!Number.isNaN(datumNumber) && datumNumber !== Infinity) {
          return datumNumber;
        }
        const dt = dateFormats.reduce((matchingDate: dayjs.Dayjs | null, currentDateFormat: string) => {
          if (matchingDate) {
            return matchingDate;
          }
          const dayjsObj = dayjs(datum as string, currentDateFormat, true);
          if (dayjsObj.isValid()) {
            return dayjsObj;
          }
          return null;
        }, null);
        if (dt) {
          return new Date(dt.valueOf());
        }
        const special = SpecialValues[datum as string];
        if (special !== undefined) {
          return special;
        }
        return datum;
      };

    // Configure uDSV schema
    const schemaOpts: Record<string, unknown> = {};
    if (opts.parserOptions?.delimiter) {
      schemaOpts.col = opts.parserOptions.delimiter;
    }
    if (opts.parserOptions?.quote) {
      schemaOpts.quote = opts.parserOptions.quote;
    }

    // Infer schema and parse
    const schema = inferSchema(csvString, schemaOpts);
    // Don't treat first row as header - we want all rows as data
    schema.skip = 0;
    const parser = initParser(schema);
    
    // Parse to string arrays (we'll do our own type conversion with map)
    const rows = parser.stringArrs(csvString);

    // Add rows to worksheet
    for (const row of rows) {
      worksheet.addRow(row.map(map));
    }

    return worksheet;
  }

  /**
   * @deprecated since version 4.0. You should use `CSV#read` instead.
   */
  createInputStream(): never {
    throw new Error(
      '`CSV#createInputStream` is deprecated. You should use `CSV#read` instead. This method will be removed in version 5.0.'
    );
  }

  /**
   * Write worksheet to a writable target
   */
  async write(
    target: { write: (data: string) => void; end: () => void },
    options?: CSVWriteOptions
  ): Promise<void> {
    const opts = options || {};
    const formatterOpts = opts.formatterOptions || {};
    const delimiter = formatterOpts.delimiter || ',';
    const quote = formatterOpts.quote || '"';
    const escape = formatterOpts.escape || '"';
    const rowDelimiter = formatterOpts.rowDelimiter || '\n';

    const workbookRecord = this.workbook as Record<string, unknown>;
    const worksheet = (workbookRecord.getWorksheet as (id?: string | number) => Worksheet | undefined)(
      opts.sheetName || opts.sheetId
    );

    const {dateFormat, dateUTC} = opts;
    const map =
      opts.map ||
      ((value: unknown): unknown => {
        if (value) {
          const valueRecord = value as Record<string, unknown>;
          if (valueRecord.text || valueRecord.hyperlink) {
            return valueRecord.hyperlink || valueRecord.text || '';
          }
          if (valueRecord.formula || valueRecord.result) {
            return valueRecord.result || '';
          }
          if (value instanceof Date) {
            if (dateFormat) {
              return dateUTC
                ? dayjs.utc(value).format(dateFormat)
                : dayjs(value).format(dateFormat);
            }
            return dateUTC ? dayjs.utc(value).format() : dayjs(value).format();
          }
          if (valueRecord.error) {
            return valueRecord.error;
          }
          if (typeof value === 'object') {
            return JSON.stringify(value);
          }
        }
        return value;
      });

    const includeEmptyRows = opts.includeEmptyRows === undefined || opts.includeEmptyRows;
    let lastRow = 1;

    if (worksheet) {
      const worksheetRecord = worksheet as Record<string, unknown>;
      (worksheetRecord.eachRow as (callback: (row: { values: unknown[] }, rowNumber: number) => void) => void)(
        (row, rowNumber) => {
          if (includeEmptyRows) {
            while (lastRow++ < rowNumber - 1) {
              target.write(rowDelimiter);
            }
          }
          const values = [...row.values];
          values.shift(); // Remove first undefined element
          const csvRow = formatCSVRow(values.map(map), delimiter, quote, escape);
          target.write(csvRow + rowDelimiter);
          lastRow = rowNumber;
        }
      );
    }

    target.end();
  }

  async writeFile(filename: string, options?: CSVFileWriteOptions): Promise<void> {
    const buffer = await this.writeBuffer(options);
    await Bun.write(filename, buffer);
  }

  async writeBuffer(options?: CSVWriteOptions): Promise<Buffer> {
    const chunks: string[] = [];
    const target = {
      write: (data: string) => chunks.push(data),
      end: () => {},
    };
    await this.write(target, options);
    return Buffer.from(chunks.join(''));
  }
}

export default CSV;
