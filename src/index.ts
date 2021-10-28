import fs from 'fs';
import xlsx from 'xlsx';

import * as r from './lib/request';
import * as utils from './lib/utils';

import { Row, Table } from './types';

/**
 * Apply `transform` to each cell in `sheet`
 *
 * @param sheet - workbook sheet
 * @param transform - function applied to each cell
 * @returns workbook sheet with transformed cell values
 */
function transformCells(
  sheet: xlsx.WorkSheet,
  transform = (cell: xlsx.CellObject): xlsx.CellObject => cell
): xlsx.WorkSheet {
  if (!sheet['!ref']) return sheet;
  const range = xlsx.utils.decode_range(sheet['!ref']);
  for (let rowIdx = range.s.r; rowIdx <= range.e.r; rowIdx++) {
    const row = xlsx.utils.encode_row(rowIdx);
    for (let colIdx = range.s.c; colIdx <= range.e.c; colIdx++) {
      const col = xlsx.utils.encode_col(colIdx);
      const cell = sheet[col + row];
      sheet[col + row] = transform(cell);
    }
  }
  return sheet;
}

/**
 * Split table into header and body
 *
 * Start and end of the data body are the first and
 * last row in table that are complete
 *
 * @param table - 2d matrix
 * @returns - two tables containing data and the head of the table resp.
 */
function split<T>(table: Table<T>): { data: Table<T>; header: Table<T> } {
  if (table.length === 0) return { header: [], data: [] };

  const nRows = table.length;
  const nCols = table[0].length;

  const findStartIndex = (table: Table<T>): number | null => {
    const idx = table.findIndex(
      (row) => row.filter((cell) => cell).length === nCols
    );
    return idx >= 0 ? idx : null;
  };

  // the first complete row marks the beginning of the data
  const firstRow = findStartIndex(table) || 0;

  // the last complete row marks the end of the data
  const lastRow = nRows - 1 - (findStartIndex([...table].reverse()) || 0);

  const header = table.slice(0, firstRow);
  const data = table.slice(firstRow, lastRow + 1);

  return { header, data };
}

async function main() {
  // ask for excel file
  const filename = await r.requestExcelFilename();

  // get sheets to process
  const workbook = xlsx.readFile(filename);
  const { SheetNames: sheetNames } = workbook;
  const selectedSheets = await r.requestSheets(sheetNames);

  // number format
  const isGermanFormat = await r.requestFormat();

  let prevColNames;
  for (let sheetNum = 0; sheetNum < selectedSheets.length; sheetNum++) {
    const sheetName = selectedSheets[sheetNum];

    // grab sheet data from excel
    let sheet = workbook.Sheets[sheetName];

    sheet = transformCells(sheet, (cell) => {
      if (!cell) return cell;
      if (cell.t === 's') {
        // necessary to ensure cells do not contains new line characters
        if (cell.w) cell.w = utils.removeLineBreaks(cell.w);
        if (cell.v)
          cell.v =
            typeof cell.v === 'string'
              ? utils.removeLineBreaks(cell.v)
              : cell.v;
      } else if (isGermanFormat && cell.t === 'n') {
        // convert to English formatting style
        if (cell.w) {
          cell.w = cell.w
            // @ is just a temporary placeholder
            .replaceAll(',', '@')
            .replaceAll('.', ',')
            .replaceAll('@', '.');
          if (cell.v) cell.v = +cell.w;
        }
      }

      return cell;
    });

    // convert to csv
    const csv = xlsx.utils.sheet_to_csv(sheet, {
      rawNumbers: true,
      blankrows: false,
    });

    // tabular data as 2d matrix
    const table = csv.split(/\r\n|\r|\n/).map(utils.toRow);

    // separate data from meta information
    const splitTable = split(table);
    const { header } = splitTable;
    let { data } = splitTable;

    // get column names from the user
    const color = sheetNum % 2 === 0 ? 'green' : 'yellow';
    const colNames = (await r.requestColumnNames(
      sheetName,
      header,
      color,
      prevColNames
    )) as Array<string>;
    prevColNames = colNames;

    // add column names to data and remove columns to ignore
    data.unshift(colNames);
    data = utils.transpose(
      utils
        .transpose(data)
        .filter((row: Row<string>): boolean => !row[0].includes('ignored'))
    );

    // save tabular data to csv file
    const out = await r.requestOutFile(filename, sheetName);
    fs.writeFileSync(out, utils.toCsv(data));
  }
}

(async () => {
  await main();
})();
