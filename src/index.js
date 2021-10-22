const inquirer = require('inquirer');

inquirer.registerPrompt('fuzzypath', require('inquirer-fuzzy-path'));

const handleError = (error) => {
  if (error.isTtyError) {
    // Prompt couldn't be rendered in the current environment
  } else {
    // Something else went wrong
  }
};

inquirer
  .prompt([
    {
      type: 'fuzzypath',
      name: 'filename',
      excludeFilter: (path) => !path.endsWith('.xlsx'),
      itemType: 'any',
      message: 'Select file',
    },
  ])
  .then((answers) => {
    console.log(answers);
  })
  .catch(handleError);
