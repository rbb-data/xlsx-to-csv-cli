import path from 'path';
import inquirer from 'inquirer';
import inquirerFuzzyPath from 'inquirer-fuzzy-path';
import chalk, { ForegroundColor } from 'chalk';

import * as utils from './utils';

import { Row, Table } from '../types';

inquirer.registerPrompt('fuzzypath', inquirerFuzzyPath);

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
 * Request name of Excel file from the user
 *
 * @param extension - allowed filename extension
 * @returns filename
 */
export async function requestExcelFilename(
  extension = 'xlsx'
): Promise<string> {
  const { filename } = (await prompt([
    {
      type: 'fuzzypath',
      name: 'filename',
      message: 'Select file',
      itemType: 'file',
      suffix: ` (*.${extension})`,
      excludeFilter: (path: string): boolean => !path.endsWith(`.${extension}`),
    },
  ])) as { filename: string };
  return filename;
}

/**
 * Request sheets to process from the user
 *
 * @param sheetNames - available sheets
 * @returns selected sheets
 */
export async function requestSheets(
  sheetNames: Array<string>
): Promise<Array<string>> {
  const { selectedSheets } = (await prompt([
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
  ])) as { selectedSheets: Array<string> };
  return selectedSheets;
}

/**
 * Check the number format
 *
 * @returns true if number should be formatted
 */
export async function requestFormat(): Promise<boolean> {
  const { isGermanFormat } = (await prompt([
    {
      type: 'confirm',
      name: 'isGermanFormat',
      message:
        'Are numbers formatted in German and do you want them to be converted to English-style numbers?',
      default: false,
    },
  ])) as { isGermanFormat: boolean };
  return isGermanFormat;
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
export async function requestColumnNames(
  sheetName: string,
  header: Table<string>,
  color: typeof ForegroundColor = 'green',
  prevColNames?: Array<string>
): Promise<Array<string>> {
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

/**
 * Request name of a result file from user for a specific sheet
 *
 * @param filename name of the original Excel file
 * @param sheetName name of the sheet
 * @param color message color
 * @returns name of the result file
 */
export async function requestOutFile(
  filename: string,
  sheetName: string,
  color: typeof ForegroundColor = 'green'
): Promise<string> {
  const c = (text: string): string => chalk[color](text);

  const suffix = sheetName
    .replace('/', '-')
    .replace(' - ', '-')
    .replace(' ', '-');

  const { out } = (await prompt([
    {
      type: 'input',
      name: 'out',
      message: c(sheetName + ':') + ' Name of result file',
      default: `${filename.replace(path.extname(filename), '')}_${suffix}.csv`,
      prefix: c('?'),
    },
  ])) as { out: string };
  return out;
}
