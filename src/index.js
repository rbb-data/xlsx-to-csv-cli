const fs = require('fs');
const inquirer = require('inquirer');
const inquirerFuzzyPath = require('inquirer-fuzzy-path');
const xlsx = require('xlsx');

const utils = require('./lib/utils');

inquirer.registerPrompt('fuzzypath', inquirerFuzzyPath);

const EXTENSION = '.xlsx';

function prompt(questions) {
  try {
    return inquirer.prompt(questions);
  } catch (error) {
    if (error.isTtyError) {
      // Prompt couldn't be rendered in the current environment
    } else {
      // Something else went wrong
    }
  }
}

function split(table) {
  const nRows = table.length;
  const nCols = table[0].length;

  const findStartIndex = (table) => {
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

async function requestColumnNames(sheetName, header, prevColNames) {
  // reuse column names from the previous table?
  const { usePrevColNames } = await prompt([
    {
      type: 'confirm',
      name: 'usePrevColNames',
      message: `${sheetName}: Do you want to reuse the column names you assigned to the previous table?`,
      default: false,
      when() {
        return prevColNames;
      },
    },
  ]);

  if (usePrevColNames) return prevColNames;

  // ask for column names
  const requests = utils.transpose(header).map((heading, j) => ({
    type: 'input',
    name: `colName-${j}`,
    message: `${sheetName}: Name of column #${String(j + 1).padStart(2, '0')}`,
    default: heading
      .filter((cell) => cell)
      .map((cell) => cell.replace('\r', '').trim())
      .join(' / '),
  }));

  // transform answers into column names
  const answers = await prompt(requests);
  const indexedAnswers = Object.entries(answers).map(([key, value]) => [
    +key.replace('colName-', ''),
    value,
  ]);
  indexedAnswers.sort((a, b) => a[0] - b[0]);
  const colNames = indexedAnswers.map(([, value]) => value);

  return colNames;
}

async function main() {
  // ask for excel file
  const { filename } = await prompt([
    {
      type: 'fuzzypath',
      name: 'filename',
      message: 'Select file',
      itemType: 'file',
      suffix: ` (*${EXTENSION})`,
      excludeFilter: (path) => !path.endsWith(EXTENSION),
    },
  ]);

  // read sheet names
  const workbook = xlsx.readFile(filename);
  const { SheetNames: sheetNames } = workbook;

  // ask the user which sheets to convert
  const { selectedSheets } = await prompt([
    {
      type: 'checkbox',
      name: 'selectedSheets',
      message: 'Select sheets',
      choices: sheetNames,
      default: sheetNames,
      loop: false,
    },
  ]);

  let prevColNames;
  for (let sheetNum = 0; sheetNum < selectedSheets.length; sheetNum++) {
    const sheetName = selectedSheets[sheetNum];

    // grab sheet data from excel
    const csv = xlsx.utils.sheet_to_csv(workbook.Sheets[sheetName]);

    // tabular data
    let table = csv.split('\n').map(utils.toRow);

    // remove empty rows and cols
    table = table.filter(utils.hasEntry);
    table = utils.transpose(utils.transpose(table).filter(utils.hasEntry));

    // separate data from meta information
    const { header, data } = split(table);

    // get column names from the user
    const colNames = await requestColumnNames(sheetName, header, prevColNames);
    prevColNames = colNames;
    data.unshift(colNames);

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
