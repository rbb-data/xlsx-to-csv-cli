import fs from 'fs';
import inquirer from 'inquirer';
import inquirerFuzzyPath from 'inquirer-fuzzy-path';
import chalk from 'chalk';

import * as utils from './utils';

import { Row, Table, Question } from '../types';

inquirer.registerPrompt('fuzzypath', inquirerFuzzyPath);

// TODO: update doc
// TODO: check for unnecessary type annotations

/**
 * Asks the user a collection of questions
 *
 * @param questions - collection of questions to ask the user
 * @return the user's answers
 */
function prompt(questions: inquirer.QuestionCollection) {
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
 * Ask the user to confirm a single question
 *
 * @param question - ask the user to confirm
 * @returns answer from user
 */
export async function confirm(question: Question) {
  const { isConfirmed } = (await prompt([
    {
      type: 'confirm',
      name: 'isConfirmed',
      message: 'Do you confirm?',
      default: false,
      ...question,
    },
  ])) as { isConfirmed: boolean };
  return isConfirmed;
}

/**
 * Request input from the user
 *
 * @param question - request input from the user
 * @returns the user's input
 */
export async function requestString(question: Question) {
  const { string } = (await prompt([
    {
      type: 'input',
      name: 'string',
      message: 'Enter string',
      ...question,
    },
  ])) as { string: string };
  return string;
}

/**
 * Request a filename with `extension` from the user
 *
 * @param extension - allowed extension
 * @param question - additional question options
 * @returns path
 */
export async function requestFile(extension: string, question?: Question) {
  const { filename } = (await prompt([
    {
      type: 'fuzzypath',
      name: 'filename',
      message: 'Select file',
      itemType: 'file',
      suffix: ` (*.${extension})`,
      excludeFilter: (path: string): boolean => !path.endsWith(`.${extension}`),
      ...question,
    },
  ])) as { filename: string };
  return filename;
}

/**
 * Request configuration from the user
 *
 * @returns configuration
 */
export async function requestConfig() {
  const useConfig = await confirm({
    message: 'Do you want to use an external configuration to pre-fill fields?',
  });
  if (!useConfig) return null;

  const filename = await requestFile('json', { message: 'Select config file' });
  const config = JSON.parse(fs.readFileSync(filename, 'utf-8'));

  return config;
}

/**
 * Request sheets to process from the user
 *
 * @param sheetNames - available sheets
 * @param defaultSheets - subset of `sheetNames` selected by default
 * @returns selected sheets
 */
export async function requestSheets(
  sheetNames: Array<string>,
  defaultSheets: Array<string>
) {
  const { selectedSheets } = (await prompt([
    {
      type: 'checkbox',
      name: 'selectedSheets',
      message: 'Select sheets',
      choices: sheetNames,
      default: defaultSheets || sheetNames,
      loop: false,
      validate(selectedSheets) {
        return selectedSheets.length > 0 ? true : 'Select at least one sheet';
      },
    },
  ])) as { selectedSheets: Array<string> };
  return selectedSheets;
}

/**
 * Ask the user to specify column names for a sheet
 *
 * @param sheetName - name of the sheet
 * @param header - table with headings
 * @param color - to highlight sheet name
 * @param defaultColNames - column names selected by default
 * @param prevColNames - previously chosen column names
 * @returns column names
 */
export async function requestColumnNames(
  sheetName: string,
  header: Table<string>,
  color: (text: string) => string = chalk.green,
  defaultColNames: Array<string>,
  prevColNames?: Array<string>
) {
  console.log();

  // reuse column names from the previous table?
  const usePrevColNames = await confirm({
    message: color(
      'Do you want to reuse the column names you assigned to the previous table?'
    ),
    prefix: color('?'),
    when: () => !defaultColNames && prevColNames,
  });

  if (usePrevColNames && prevColNames) return prevColNames;

  // ask for column names
  const requests = utils
    .transpose(header)
    .map((heading: Row<string>, j: number) => {
      const defaultValue = defaultColNames
        ? defaultColNames[j]
        : heading
            .filter((cell: string): string => cell)
            .map((cell: string): string => cell.replace('\r', '').trim())
            .join(' / ');

      return {
        type: 'input',
        name: `colName-${j}`,
        message:
          color(sheetName + ': ') +
          `Name of column #${String(j + 1).padStart(2, '0')}`,
        default: defaultValue || `col_${j + 1}`,
        prefix: color('?'),
        suffix: ' (Type "-" to ignore)',
        filter: (colName: string) =>
          colName.toLowerCase() === '-' ? chalk.dim.italic('ignored') : colName,
      };
    });

  // transform answers into column names
  const answers = (await prompt(requests)) as Record<string, string>;
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
