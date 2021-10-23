const isSpecial = (char) => ['"', "'", ','].includes(char);
const isQuote = (char) => ['"', "'"].includes(char);
const isComma = (char) => char === ',';

function toRow(line) {
  let row = [];
  let idx = 0;
  let withinQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const char = line[i];

    if (isSpecial(char)) {
      if (isQuote(char)) withinQuotes = !withinQuotes;
      else if (isComma(char) && !withinQuotes) {
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

function transpose(table) {
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

function hasEntry(row) {
  return row.some((cell) => cell);
}

function toCsv(table) {
  return table
    .map((row) => row.map((cell) => `"${cell}"`).join(','))
    .join('\n');
}

module.exports = {
  isSpecial,
  isQuote,
  isComma,
  toRow,
  transpose,
  hasEntry,
  toCsv,
};
