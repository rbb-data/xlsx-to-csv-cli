import path from 'path';
import inquirer from 'inquirer';
import inquirerFuzzyPath from 'inquirer-fuzzy-path';
import chalk, { ForegroundColor } from 'chalk';

import * as utils from './utils';

import { Row, Table, Config, ExtendedConfig } from '../types';

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

type RequestOptions =
  | RequestColumnNamesOptions
  | RequestOutFileOptions
  | RequestSheetsOptions
  | RequestFileOptions
  | RequestOutConfigFileOptions
  | RequestFormatOptions;

/**
 * Update `config` with the user's answers stored in `key`
 *
 * @param key - name of the requested variable
 * @param config - current configuration
 * @returns updated configuration
 */
export default async function request(
  key: keyof Config,
  config?: ExtendedConfig,
  options?: RequestOptions
): Promise<ExtendedConfig> {
  if (!config) config = {};
  if (!options && key === 'filename')
    options = { extension: 'xlsx' } as RequestFileOptions;

  const _request = {
    filename: requestFile,
    configFilename: requestConfigFilename,
    sheets: requestSheets,
    isGermanFormat: requestFormat,
    colNames: requestColumnNames,
    out: requestOutFile,
    configOut: requestOutConfigFile,
  }[key];

  return {
    ...config,
    // in a perfect world, one would check for the correct type here
    [key]: await _request(options as any),
  };
}

interface RequestFileOptions {
  extension: string;
  message: string;
}

/**
 * Request a filename with `extension` from the user
 *
 * @param extension - allowed extension
 * @param message - message to display
 * @returns path
 */
async function requestFile({
  extension,
  message = 'Select file',
}: RequestFileOptions): Promise<string> {
  const { filename } = (await prompt([
    {
      type: 'fuzzypath',
      name: 'filename',
      message,
      itemType: 'file',
      suffix: ` (*.${extension})`,
      excludeFilter: (path: string): boolean => !path.endsWith(`.${extension}`),
    },
  ])) as { filename: string };
  return filename;
}

/**
 * Request configuration from the user
 *
 * @returns configuration
 */
async function requestConfigFilename(): Promise<string | null> {
  const { useConfig } = (await prompt([
    {
      type: 'confirm',
      name: 'useConfig',
      message:
        'Do you want to use an external configuration to pre-fill fields?',
      default: false,
    },
  ])) as { useConfig: boolean };

  if (!useConfig) return null;

  const filename = await requestFile({
    extension: 'json',
    message: 'Select config file',
  });

  return filename;
}

interface RequestSheetsOptions {
  sheetNames: Array<string>;
  defaultSheets?: Array<string>;
}

/**
 * Request sheets to process from the user
 *
 * @param sheetNames - available sheets
 * @returns selected sheets
 */
async function requestSheets({
  sheetNames,
  defaultSheets,
}: RequestSheetsOptions): Promise<Array<string>> {
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

interface RequestFormatOptions {
  defaultValue: boolean;
}

/**
 * Check the number format
 *
 * @returns true if number should be formatted
 */
async function requestFormat({
  defaultValue = false,
}: RequestFormatOptions): Promise<boolean> {
  const { isGermanFormat } = (await prompt([
    {
      type: 'confirm',
      name: 'isGermanFormat',
      message:
        'Are numbers formatted in German and do you want them to be converted to English-style numbers?',
      default: defaultValue,
    },
  ])) as { isGermanFormat: boolean };
  return isGermanFormat;
}

interface RequestColumnNamesOptions {
  sheetName: string;
  header: Table<string>;
  color: typeof ForegroundColor;
  prevColNames?: Array<string>;
  defaultColNames?: Array<string>;
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
async function requestColumnNames({
  sheetName,
  header,
  color = 'green',
  prevColNames,
  defaultColNames,
}: RequestColumnNamesOptions): Promise<Array<string>> {
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
        return !defaultColNames && prevColNames;
      },
    },
  ])) as { usePrevColNames: boolean };

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
          c(sheetName + ': ') +
          `Name of column #${String(j + 1).padStart(2, '0')}`,
        default: defaultValue,
        prefix: c('?'),
        suffix: ' (Type "no" to ignore)',
        filter: (colName: string): string =>
          colName.toLowerCase() === 'no'
            ? chalk.dim.italic('ignored')
            : colName,
      };
    });

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

interface RequestOutFileOptions {
  filename: string;
  sheetName: string;
  color: typeof ForegroundColor;
  defaultFilename?: string;
}

/**
 * Request name of a result file from user for a specific sheet
 *
 * @param filename name of the original Excel file
 * @param sheetName name of the sheet
 * @param color message color
 * @returns name of the result file
 */
async function requestOutFile({
  filename,
  sheetName,
  color = 'green',
  defaultFilename,
}: RequestOutFileOptions): Promise<string> {
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
      default:
        defaultFilename ||
        `${filename.replace(path.extname(filename), '')}_${suffix}.csv`,
      prefix: c('?'),
    },
  ])) as { out: string };
  return out;
}

interface RequestOutConfigFileOptions {
  filename: string;
  defaultFilename?: string;
}

async function requestOutConfigFile({
  filename,
  defaultFilename,
}: RequestOutConfigFileOptions): Promise<string> {
  console.log();

  const { out } = (await prompt([
    {
      type: 'confirm',
      name: 'saveConfig',
      message: 'Do you want to save the specified configuration?',
      default: false,
    },
    {
      type: 'input',
      name: 'out',
      message: 'Name of the config file',
      default:
        defaultFilename ||
        `${filename.replace(path.extname(filename), '')}.json`,
      when: ({ saveConfig }) => saveConfig,
    },
  ])) as { saveConfig: boolean; out: string };

  return out;
}
