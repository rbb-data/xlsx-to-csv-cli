const fs = require('fs');
const inquirer = require('inquirer');
const xlsx = require('xlsx');

inquirer.registerPrompt('fuzzypath', require('inquirer-fuzzy-path'));

const EXTENSION = '.xlsx';

async function main() {
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

  // ask for excel file
  const { filename } = await prompt([
    {
      type: 'fuzzypath',
      name: 'filename',
      message: 'Select file',
      itemType: 'any',
      suffix: `*${EXTENSION}`,
      excludeFilter: (path) => !path.endsWith(EXTENSION),
    },
  ]);

  // read sheet names
  const workbook = xlsx.readFile(filename);
  const { SheetNames: sheetNames } = workbook;

  // ask the user which sheets to convert
  const { sheets } = await prompt([
    {
      type: 'checkbox',
      name: 'sheets',
      message: 'Select sheets',
      choices: sheetNames,
      default: sheetNames,
      loop: false,
    },
  ]);

  // write sheets to csv files
  sheets.forEach((sheetName) => {
    const csv = xlsx.utils.sheet_to_csv(workbook.Sheets[sheetName]);
    const suffix = sheetName
      .replace('/', '-')
      .replace(' - ', '-')
      .replace(' ', '-');
    const out = `${filename.replace(EXTENSION, '')}_${suffix}.csv`;
    fs.writeFileSync(out, csv);
  });
}

(async () => {
  await main();
})();
