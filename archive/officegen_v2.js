const fs = require('fs');
const readline = require('readline');
const officegen = require('officegen');
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
  const pptx = officegen('pptx');

  function addSlideFromTemplate(templatePath) {
    const slide = pptx.makeNewSlide();

    const templateContent = fs.readFileSync(templatePath);
    slide.addStream(templateContent);

    return slide;
  }

  excelData.forEach((row) => {
    const templateNumber = row['Template'];

    let templatePath;
    switch (templateNumber) {
      case 1:
        templatePath = 'templates/template_1.pptx';
        break;
      case 2:
        templatePath = 'templates/template_2.pptx';
        break;
      case 3:
        templatePath = 'templates/template_3.pptx';
        break;
      default:
        console.error('Invalid template number in Excel data.');
        rl.close();
        return;
    }

    addSlideFromTemplate(templatePath);
  });

  const date = new Date().toISOString().split('T')[0];
  const outputFileName = `${date}_${projectName}_${version}.pptx`;

  const outputStream = fs.createWriteStream(outputFileName);
  pptx.generate(outputStream);

  outputStream.on('finish', () => {
    console.log(`PowerPoint presentation "${outputFileName}" has been created.`);
    rl.close();
  });

  outputStream.on('error', (err) => {
    console.error('Error generating PowerPoint presentation:', err);
    rl.close();
  });
});
