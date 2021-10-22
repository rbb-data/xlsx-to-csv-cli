const inquirer = require('inquirer');
const xlsx = require('xlsx');

inquirer.registerPrompt('fuzzypath', require('inquirer-fuzzy-path'));

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
      excludeFilter: (path) => !path.endsWith('.xlsx'),
    },
  ]);

  // read sheet names
  const workbook = xlsx.readFile(filename);
  const { SheetNames: sheetNames } = workbook;

  // ask for sheets to be converted
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

  console.log(sheets);
}

(async () => {
  await main();
})();
