// @ts-check

/**
 * Is `char` a special character
 *
 * @param {string} char - Single character
 * @returns {boolean}
 */
const isSpecial = (char) => ['"', "'", ','].includes(char);

/**
 * Is `char` a quotation mark
 *
 * @param {string} char - Single character
 * @returns {boolean}
 */
const isQuotationMark = (char) => ['"', "'"].includes(char);

/**
 * Is `char` a csv separator
 *
 * @param {string} char - Single character
 * @returns {boolean}
 */
const isSeparator = (char) => char === ',';

/**
 * Convert line to csv row
 *
 * @param {string} line
 * @returns {Array<string>} row of cells
 */
function toRow(line) {
  let row = [];
  let idx = 0;
  let withinQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const char = line[i];

    if (isSpecial(char)) {
      if (isQuotationMark(char)) withinQuotes = !withinQuotes;
      else if (isSeparator(char) && !withinQuotes) {
        if (!row[idx]) row[idx] = '';
        idx += 1;
      }
      continue;
    }

    if (!row[idx]) row[idx] = '';
    row[idx] += char;
  }
  return row;
}

/**
 * Transpose matrix
 *
 * @template T
 * @param {import("../types").Table<T>} table - of size n x m
 * @returns {import("../types").Table<T>} Transposed table of size m x n
 */
function transpose(table) {
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
 * @param {import("../types").Row<any>} row
 * @returns {boolean}
 */
function hasEntry(row) {
  return row.some((cell) => cell);
}

/**
 * Convert data table into CSV
 *
 * @param {import("../types").Table<any>} table
 * @returns {string} - CSV-formatted string
 */
function toCsv(table) {
  return table
    .map((row) => row.map((cell) => `"${cell}"`).join(','))
    .join('\n');
}

/**
 * Remove line breaks from the given string
 *
 * @param {string} str
 * @returns {string} Single-line string
 */
function removeLineBreaks(str) {
  // @ts-ignore
  return str.replaceAll('\n', ' ').replaceAll('\r', '');
}

module.exports = {
  isSpecial,
  isQuote: isQuotationMark,
  isComma: isSeparator,
  toRow,
  transpose,
  hasEntry,
  toCsv,
  removeLineBreaks,
};
