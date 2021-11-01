import path from 'path';

import { Row, Table } from '../types';

/**
 * Is `char` a special character in CSVs (`""` or `,`)
 *
 * @param char - single character
 * @returns true if `char` is a quotation mark or comma
 */
export const isSpecial = (char: string): boolean =>
  isQuotationMark(char) || isSeparator(char);

/**
 * Is `char` a quotation mark
 *
 * @param char - single character
 * @returns true if `char` is a quotation mark
 */
export const isQuotationMark = (char: string): boolean => char === '"';

/**
 * Is `char` a csv separator
 *
 * @param char - single character
 * @returns true if `char` is a comma
 */
export const isSeparator = (char: string): boolean => char === ',';

/**
 * Convert line to csv row
 *
 * @param line - single line in csv format
 * @returns row of cells
 */
export function toRow(line: string): Row<string> {
  const row = [];
  let idx = 0;
  let withinQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const char = line[i];

    if (isSpecial(char)) {
      if (isQuotationMark(char)) {
        withinQuotes = !withinQuotes;
        continue;
      }

      if (isSeparator(char) && !withinQuotes) {
        if (!row[idx]) row[idx] = '';
        idx += 1;
        continue;
      }
    }

    if (!row[idx]) row[idx] = '';
    row[idx] += char;
  }
  return row;
}

/**
 * Transpose matrix
 *
 * @param table - of size n x m
 * @returns Transposed `table` of size m x n
 */
export function transpose<T>(table: Table<T>): Table<T> {
  if (table.length === 0) return table;

  const nRows = table.length;
  const nCols = table[0].length;

  const transposed = Array.from(Array(nCols), () => new Array(nRows));
  for (let i = 0; i < nRows; i++) {
    for (let j = 0; j < nCols; j++) {
      transposed[j][i] = table[i][j];
    }
  }

  return transposed;
}

/**
 * Has at least one truthy value
 *
 * @param row - csv row
 * @returns false if all values in `row` are falsy
 */
export function hasEntry(row: Row<string>): boolean {
  return row.some((cell) => cell);
}

/**
 * Convert data table into CSV
 *
 * @param table - 2d matrix
 * @returns csv-formatted string
 */
export function toCsv(table: Table<string>): string {
  return table
    .map((row) => row.map((cell) => `"${cell}"`).join(','))
    .join('\n');
}

/**
 * Remove line breaks from the given string
 *
 * @param str - possibly multi-line string
 * @returns Single-line string
 */
export function removeLineBreaks(str: string): string {
  return str.replaceAll('\n', ' ').replaceAll('\r', '');
}

/**
 * Replace extension of `filename` with `suffix`
 *
 * @param filename - path to file
 * @param suffix - suffix to add to `filename`
 * @returns `filename` with `suffix` appended
 */
export function replaceExtension(filename: string, suffix: string) {
  return `${filename.replace(path.extname(filename), '')}${suffix}`;
}

/**
 * Convert German to English number format
 *
 * @param str - string containing numbers in German format
 * @returns string containing numbers in English format
 */
export function toEnglishFormat(str: string) {
  // @ is just a temporary placeholder
  return str.replaceAll(',', '@').replaceAll('.', ',').replaceAll('@', '.');
}

/**
 * Remove characters problematic for file paths
 *
 * @param str - any string
 * @returns string that can be used as path
 */
export function normalize(str: string) {
  return str.replaceAll('/', '-').replaceAll(' - ', '-').replaceAll(' ', '-');
}
