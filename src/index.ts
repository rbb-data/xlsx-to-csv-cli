import fs from 'fs';
import xlsx from 'xlsx';
import chalk from 'chalk';

import {
  requestFile,
  requestConfig,
  requestSheets,
  confirm,
  requestColumnNames,
  requestString,
} from './lib/request';
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
function split(table: Table<string>): {
  data: Table<string>;
  header: Table<string>;
} {
  if (table.length === 0) return { header: [], data: [] };

  const nRows = table.length;
  const nCols = table[0].length;

  const findStartIndex = (
    table: Table<string>,
    requireNumber = false
  ): number | null => {
    const idx = table.findIndex(
      (row) =>
        row.filter((cell) => cell).length === nCols &&
        (!requireNumber ||
          row.some((cell) => !Number.isNaN(+cell.replaceAll(',', ''))))
    );
    return idx >= 0 ? idx : null;
  };

  // the first complete row marks the beginning of the data
  const firstRow = findStartIndex(table, true) || 0;

  // the last complete row marks the end of the data
  const lastRow = nRows - 1 - (findStartIndex([...table].reverse()) || 0);

  let header = table.slice(0, firstRow);
  let data = table.slice(firstRow, lastRow + 1);

  if (header.length === 0) header = [new Array(nCols).fill('')];
  if (data.length === 0) data = [new Array(nCols).fill('')];

  return { header, data };
}

async function main() {
  // ask for excel file
  const filename = await requestFile('xlsx');

  // ask for config file
  let config = (await requestConfig()) || {};

  // get sheets to process
  const workbook = xlsx.readFile(filename);
  const { SheetNames: sheetNames } = workbook;
  const sheets = await requestSheets(sheetNames, config.sheets);

  // number format
  const isGermanFormat = await confirm({
    message:
      'Are numbers formatted in German and do you want them to be converted to English-style numbers?',
    default: config.isGermanFormat || false,
  });

  let prevColNames;
  const colNames: Record<string, Array<string>> = {};
  const out: Record<string, string> = {};
  for (let sheetNum = 0; sheetNum < sheets.length; sheetNum++) {
    const sheetName = sheets[sheetNum];
    const color = chalk[sheetNum % 2 === 0 ? 'green' : 'yellow'];

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
          cell.w = utils.toEnglishFormat(cell.w);
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
    let table = csv.split(/\r\n|\r|\n/).map(utils.toRow);

    // remove empty rows and cols
    table = table.filter(utils.hasEntry);
    table = utils.transpose(utils.transpose(table).filter(utils.hasEntry));

    // separate data from meta information
    const splitTable = split(table);
    const { header } = splitTable;
    let { data } = splitTable;

    // get column names from the user
    colNames[sheetName] = (await requestColumnNames(
      sheetName,
      header,
      color,
      config.colNames ? config.colNames[sheetName] : undefined,
      prevColNames
    )) as Array<string>;
    prevColNames = colNames[sheetName];

    // add column names to data and remove columns to ignore
    data.unshift(colNames[sheetName]);
    data = utils.transpose(
      utils
        .transpose(data)
        .filter((row: Row<string>): boolean => !row[0].includes('ignored'))
    );

    // save to csv file
    out[sheetName] = await requestString({
      message: color(sheetName + ':') + ' Name of result file',
      default: config.out
        ? config.out[sheetName]
        : utils.replaceExtension(
            filename,
            `_${utils.normalize(sheetName)}.csv`
          ),
      prefix: color('?'),
    });
    fs.writeFileSync(out[sheetName], utils.toCsv(data));
  }

  // update config
  config = {
    ...config,
    filename,
    sheets,
    isGermanFormat,
    colNames,
    out,
  };

  // save config to file
  console.log();
  const saveConfig = await confirm({
    message: 'Do you want to save the specified configuration?',
  });
  if (saveConfig) {
    const configOut = await requestString({
      message: 'Name of the config file',
      default: utils.replaceExtension(filename, '.json'),
    });
    fs.writeFileSync(configOut, JSON.stringify(config, null, 2));
  }
}

(async () => {
  await main();
})();
