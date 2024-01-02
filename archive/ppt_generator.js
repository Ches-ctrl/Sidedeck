const fs = require('fs');
const readline = require('readline');
const pptxgen = require('pptxgenjs');
const xlsx = require('xlsx');

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

const excelFile = 'test_deck_builder.xlsx';
const workbook = xlsx.readFile(excelFile);
const worksheet = workbook.Sheets['Sheet1'];

const excelData = xlsx.utils.sheet_to_json(worksheet);
console.log('Loaded ' + excelData.length + ' rows');

rl.question('Enter the project name: ', (projectName) => {

  projectName = projectName.trim();
  const version = 'v0';

  const mainPresentation = new pptxgen();
  function copySlidesFromTemplate(templatePath) {
    const templatePresentation = new pptxgen();

    // Load the template presentation.
    // You'll need to use a library that supports loading existing PowerPoint files.
    // Example: templatePresentation.load(templatePath);

    // Iterate through the slides in the template presentation.
    // Copy each slide and add it to the main presentation.
    // Example: mainPresentation.addSlide(templatePresentation.slides[0]);
  }

  // Iterate through each row of the Excel data and add slides based on templates.
  excelData.forEach((row) => {
    const templateNumber = row['Template']; // Assuming a 'Template' column in your Excel data.

    // Determine the template file path based on the value in the Excel data.
    let templatePath;
    switch (templateNumber) {
      case 1:
        templatePath = 'templates/template1.pptx';
        break;
      case 2:
        templatePath = 'templates/template2.pptx';
        break;
      case 3:
        templatePath = 'templates/template3.pptx';
        break;
      default:
        console.error('Invalid template number in Excel data.');
        rl.close();
        return;
    }

    // Copy slides from the selected template to the main presentation.
    copySlidesFromTemplate(templatePath);
  });

  // Define the output file name.
  const date = new Date().toISOString().split('T')[0];
  const outputFileName = `${date}_${projectName}_${version}.pptx`;

  // Save the final PowerPoint document (main presentation).
  mainPresentation.save(outputFileName);

  console.log(`PowerPoint presentation "${outputFileName}" has been created.`);
  rl.close();
});
