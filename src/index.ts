import fs from 'fs';
import inquirer from 'inquirer';
import inquirerFuzzyPath from 'inquirer-fuzzy-path';
import xlsx from 'xlsx';
import chalk, { ForegroundColor } from 'chalk';

import * as utils from './lib/utils';
import { Row, Table } from './types';

inquirer.registerPrompt('fuzzypath', inquirerFuzzyPath);

const EXTENSION = '.xlsx';

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
 * Asks the user a collection of questions
 *
 * @param questions - collection of questions to ask the user
 * @return the user's answers
 */
function prompt(
  questions: inquirer.QuestionCollection
): Promise<{ [key: string]: unknown }> {
  try {
    return inquirer.prompt(questions);
  } catch (error: any) {
    if (error.isTtyError) {
      process.stderr.write(
        "Prompt couldn't be rendered in the current environment"
      );
      process.exit(1);
    } else {
      process.stderr.write('Something went wrong');
      process.exit(1);
    }
  }
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

/**
 * Ask the user to specify column names for a sheet
 *
 * @param sheetName - name of the sheet
 * @param header - table with headings
 * @param color - to highlight sheet name
 * @param prevColNames - previously chosen column names
 * @returns column names
 */
async function requestColumnNames(
  sheetName: string,
  header: Table<string>,
  color: typeof ForegroundColor,
  prevColNames?: Array<string>
): Promise<Array<string>> {
  /**
   * Add styles to text
   * @param {string} text
   * @returns {string} Colored text
   */
  const c = (text: string): string => chalk[color](text);

  console.log();

  // reuse column names from the previous table?
  const { usePrevColNames } = (await prompt([
    {
      type: 'confirm',
      name: 'usePrevColNames',
      message: c(
        'Do you want to reuse the column names you assigned to the previous table?'
      ),
      default: false,
      prefix: c('?'),
      when() {
        return prevColNames;
      },
    },
  ])) as { usePrevColNames: boolean };

  if (usePrevColNames && prevColNames) return prevColNames;

  // ask for column names
  const requests = utils
    .transpose(header)
    .map((heading: Row<string>, j: number) => ({
      type: 'input',
      name: `colName-${j}`,
      message:
        c(sheetName + ': ') +
        `Name of column #${String(j + 1).padStart(2, '0')}`,
      default: heading
        .filter((cell: string): string => cell)
        .map((cell: string): string => cell.replace('\r', '').trim())
        .join(' / '),
      prefix: c('?'),
      suffix: ' (Type "no" to ignore)',
      filter: (colName: string): string =>
        colName.toLowerCase() === 'no' ? chalk.dim.italic('ignored') : colName,
    }));

  // transform answers into column names
  const answers = (await prompt(requests)) as { [key: string]: string };
  const indexedAnswers = Object.entries(answers).map(
    ([key, value]: [string, string]): [number, string] => [
      +key.replace('colName-', ''),
      value,
    ]
  );
  indexedAnswers.sort(
    (a: [number, string], b: [number, string]) => a[0] - b[0]
  );
  const colNames = indexedAnswers.map(([, value]) => value);

  return colNames;
}

async function main() {
  // ask for excel file
  const { filename } = (await prompt([
    {
      type: 'fuzzypath',
      name: 'filename',
      message: 'Select file',
      itemType: 'file',
      suffix: ` (*${EXTENSION})`,
      excludeFilter: (path: string): boolean => !path.endsWith(EXTENSION),
    },
  ])) as { filename: string };

  // read sheet names
  const workbook = xlsx.readFile(filename);
  const { SheetNames: sheetNames } = workbook;

  const { selectedSheets, isGermanFormat } = (await prompt([
    // ask the user which sheets to convert
    {
      type: 'checkbox',
      name: 'selectedSheets',
      message: 'Select sheets',
      choices: sheetNames,
      default: sheetNames,
      loop: false,
      validate(selectedSheets) {
        return selectedSheets.length > 0 ? true : 'Select at least one sheet';
      },
    },
    // check the number formatting
    {
      type: 'confirm',
      name: 'isGermanFormat',
      message:
        'Are numbers formatted in German and do you want them to be converted to English-style numbers?',
      default: false,
    },
  ])) as { selectedSheets: Array<string>; isGermanFormat: boolean };

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
    const csv = xlsx.utils.sheet_to_csv(sheet);

    // tabular data
    let table = csv.split(/\r\n|\r|\n/).map(utils.toRow);

    // remove empty rows and cols
    table = table.filter(utils.hasEntry);
    table = utils.transpose(utils.transpose(table).filter(utils.hasEntry));

    // separate data from meta information
    const splitTable = split(table);
    const { header } = splitTable;
    let { data } = splitTable;

    // get column names from the user
    const colNames = (await requestColumnNames(
      sheetName,
      header,
      sheetNum % 2 === 0 ? 'green' : 'yellow',
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
    const suffix = sheetName
      .replace('/', '-')
      .replace(' - ', '-')
      .replace(' ', '-');
    const out = `${filename.replace(EXTENSION, '')}_${suffix}.csv`;
    fs.writeFileSync(out, utils.toCsv(data));
  }
}

(async () => {
  await main();
})();
